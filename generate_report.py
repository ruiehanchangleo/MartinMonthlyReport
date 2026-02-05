#!/usr/bin/env python3
"""
XTM Monthly Report Generator

This script connects to XTM Cloud API, retrieves monthly metrics,
generates an Excel report, and prepares it for email distribution.
"""

import json
import logging
import sys
import subprocess
import platform
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Any, Callable
from functools import wraps
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('xtm_report.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


def retry_with_backoff(max_attempts=5, initial_delay=1, backoff_factor=2, max_delay=60):
    """
    Decorator that retries a function with exponential backoff.

    Args:
        max_attempts: Maximum number of retry attempts
        initial_delay: Initial delay in seconds
        backoff_factor: Multiplier for delay after each attempt
        max_delay: Maximum delay between retries in seconds
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            delay = initial_delay
            last_exception = None

            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except requests.exceptions.Timeout as e:
                    last_exception = e
                    logger.warning(f"Timeout on attempt {attempt}/{max_attempts} for {func.__name__}: {e}")
                except requests.exceptions.ConnectionError as e:
                    last_exception = e
                    logger.warning(f"Connection error on attempt {attempt}/{max_attempts} for {func.__name__}: {e}")
                except requests.exceptions.HTTPError as e:
                    last_exception = e
                    # Don't retry on 4xx errors (client errors) except 429 (rate limit)
                    if hasattr(e, 'response') and e.response is not None:
                        status_code = e.response.status_code
                        if 400 <= status_code < 500 and status_code != 429:
                            logger.error(f"Client error {status_code} for {func.__name__}, not retrying: {e}")
                            raise
                    logger.warning(f"HTTP error on attempt {attempt}/{max_attempts} for {func.__name__}: {e}")
                except requests.exceptions.RequestException as e:
                    last_exception = e
                    logger.warning(f"Request error on attempt {attempt}/{max_attempts} for {func.__name__}: {e}")
                except Exception as e:
                    # Don't retry on unexpected exceptions
                    logger.error(f"Unexpected error in {func.__name__}: {e}")
                    raise

                # If this wasn't the last attempt, wait before retrying
                if attempt < max_attempts:
                    wait_time = min(delay, max_delay)
                    logger.info(f"Retrying {func.__name__} in {wait_time} seconds...")
                    time.sleep(wait_time)
                    delay *= backoff_factor

            # All attempts failed
            logger.error(f"All {max_attempts} attempts failed for {func.__name__}")
            raise last_exception

        return wrapper
    return decorator


class XTMReportGenerator:
    """Main class for generating XTM monthly reports."""

    # Locale code to language name mapping
    LOCALE_TO_LANGUAGE = {
        'ar_AE': 'Arabic (UAE)',
        'ar_EG': 'Arabic (Egypt)',
        'ar_SA': 'Arabic (Saudi Arabia)',
        'bg_BG': 'Bulgarian',
        'ceb': 'Cebuano',
        'cs_CZ': 'Czech',
        'da_DK': 'Danish',
        'de_DE': 'German',
        'el_CY': 'Greek (Cyprus)',
        'el_GR': 'Greek',
        'en_US': 'English (US)',
        'en_GB': 'English (UK)',
        'es_ES': 'Spanish',
        'es_MX': 'Spanish (Mexico)',
        'et_EE': 'Estonian',
        'fa_IR': 'Persian',
        'fi_FI': 'Finnish',
        'fj_FJ': 'Fijian',
        'fr_FR': 'French',
        'hr_BA': 'Croatian (Bosnia)',
        'hr_HR': 'Croatian',
        'ht_HT': 'Haitian Creole',
        'hu_HU': 'Hungarian',
        'hy_AM': 'Armenian',
        'id_ID': 'Indonesian',
        'is_IS': 'Icelandic',
        'it_IT': 'Italian',
        'ja_JP': 'Japanese',
        'ka_GE': 'Georgian',
        'kk_KZ': 'Kazakh',
        'km_KH': 'Khmer',
        'ko_KR': 'Korean',
        'lo_LA': 'Lao',
        'lt_LT': 'Lithuanian',
        'lv_LV': 'Latvian',
        'mg_MG': 'Malagasy',
        'mk_MK': 'Macedonian',
        'mn_MN': 'Mongolian',
        'ms_MY': 'Malay',
        'nl_NL': 'Dutch',
        'no_NO': 'Norwegian',
        'pl_PL': 'Polish',
        'pt_BR': 'Portuguese (Brazil)',
        'pt_PT': 'Portuguese (Portugal)',
        'ro_RO': 'Romanian',
        'ru_RU': 'Russian',
        'sk_SK': 'Slovak',
        'sl_SI': 'Slovenian',
        'sm_WS': 'Samoan',
        'sq_AL': 'Albanian',
        'sr_RS': 'Serbian',
        'sv_SE': 'Swedish',
        'sw_KE': 'Swahili (Kenya)',
        'sw_TZ': 'Swahili (Tanzania)',
        'th_TH': 'Thai',
        'tl_PH': 'Tagalog',
        'to_TO': 'Tongan',
        'tr_TR': 'Turkish',
        'ty': 'Tahitian',
        'uk_UA': 'Ukrainian',
        'ur_IN': 'Urdu',
        'vi_VN': 'Vietnamese',
        'zh_CN': 'Chinese (Simplified)',
        'zh_HK': 'Chinese (Hong Kong)',
        'zh_TW': 'Chinese (Traditional)',
    }

    def __init__(self, config_path: str = "xtm_config.json", auto_send: bool = False):
        """Initialize the report generator with configuration."""
        self.config = self._load_config(config_path)
        self.auto_send = auto_send
        self.base_url = self.config['base_url']
        self.headers = {
            'Authorization': f"{self.config['auth_type']} {self.config['auth_token']}",
            'Content-Type': 'application/json'
        }
        self.report_date = datetime.now()
        # Calculate previous month for the monthly report
        first_day_current_month = self.report_date.replace(day=1)
        last_day_previous_month = first_day_current_month - timedelta(days=1)

        self.report_month = last_day_previous_month.strftime('%Y-%m')
        self.report_month_name = last_day_previous_month.strftime('%B %Y')

        # Get year-to-date range (January 1 to end of previous month)
        self.ytd_start_month = last_day_previous_month.strftime('%Y-01')
        self.ytd_end_month = self.report_month

    def _load_config(self, config_path: str) -> Dict:
        """Load configuration from JSON file."""
        try:
            config_file = Path(config_path)
            if not config_file.exists():
                raise FileNotFoundError(f"Configuration file not found: {config_path}")

            with open(config_path, 'r') as f:
                config = json.load(f)

            # Validate required configuration keys
            required_keys = ['base_url', 'auth_type', 'auth_token', 'onedrive_path', 'email_recipients']
            missing_keys = [key for key in required_keys if key not in config]
            if missing_keys:
                raise ValueError(f"Missing required config keys: {', '.join(missing_keys)}")

            # Validate auth token is not empty
            if not config.get('auth_token') or config['auth_token'].strip() == '':
                raise ValueError("Authentication token is empty")

            # Validate email recipients
            if not isinstance(config.get('email_recipients'), list) or not config['email_recipients']:
                logger.warning("No email recipients configured")

            logger.info("Configuration loaded and validated successfully")
            return config
        except Exception as e:
            logger.error(f"Failed to load config: {e}")
            raise

    def _locale_to_language_name(self, locale_code: str) -> str:
        """Convert locale code to language name."""
        return self.LOCALE_TO_LANGUAGE.get(locale_code, locale_code)

    def _run_health_checks(self) -> bool:
        """Run comprehensive health checks before starting report generation."""
        logger.info("Running health checks...")
        all_passed = True

        # 1. Check API connectivity
        try:
            logger.info("Checking API connectivity...")
            test_response = self._make_request('projects', params={'count': 1})
            logger.info("✓ API connectivity check passed")
        except Exception as e:
            logger.error(f"✗ API connectivity check failed: {e}")
            all_passed = False

        # 2. Check output directory
        try:
            logger.info("Checking output directory...")
            output_dir = Path(self.config['onedrive_path'])

            # Check if directory exists or can be created
            if not output_dir.exists():
                try:
                    output_dir.mkdir(parents=True, exist_ok=True)
                    logger.info(f"✓ Created output directory: {output_dir}")
                except Exception as e:
                    logger.warning(f"⚠ Cannot create primary output directory, will use fallback: {e}")

            # Check write permissions
            if output_dir.exists():
                test_file = output_dir / '.write_test'
                try:
                    test_file.write_text('test')
                    test_file.unlink()
                    logger.info(f"✓ Output directory is writable: {output_dir}")
                except Exception as e:
                    logger.warning(f"⚠ Output directory not writable, will use fallback: {e}")
            else:
                logger.warning(f"⚠ Output directory does not exist, will use fallback locations")

        except Exception as e:
            logger.warning(f"⚠ Output directory check failed, will use fallback: {e}")

        # 3. Check disk space (warn if less than 100MB free)
        try:
            logger.info("Checking disk space...")
            import shutil
            stat = shutil.disk_usage(Path.home())
            free_mb = stat.free / (1024 * 1024)
            if free_mb < 100:
                logger.warning(f"⚠ Low disk space: {free_mb:.0f}MB free")
            else:
                logger.info(f"✓ Sufficient disk space: {free_mb:.0f}MB free")
        except Exception as e:
            logger.warning(f"⚠ Could not check disk space: {e}")

        # 4. Check required Python packages
        try:
            logger.info("Checking required packages...")
            import openpyxl
            import requests
            logger.info("✓ All required packages are available")
        except ImportError as e:
            logger.error(f"✗ Missing required package: {e}")
            all_passed = False

        # 5. Validate date calculations
        try:
            logger.info("Validating date calculations...")
            if not self.report_month or not self.ytd_start_month:
                raise ValueError("Invalid date calculations")
            logger.info(f"✓ Report period: {self.report_month_name}")
            logger.info(f"✓ YTD period: {self.ytd_start_month} to {self.ytd_end_month}")
        except Exception as e:
            logger.error(f"✗ Date validation failed: {e}")
            all_passed = False

        if all_passed:
            logger.info("✓ All critical health checks passed")
        else:
            logger.warning("⚠ Some health checks failed, but will attempt to continue")

        return all_passed

    @retry_with_backoff(max_attempts=5, initial_delay=2, backoff_factor=2, max_delay=60)
    def _make_request(self, endpoint: str, method: str = 'GET', params: Dict = None, data: Dict = None) -> Any:
        """Make API request to XTM with error handling."""
        url = f"{self.base_url}/{endpoint}"
        try:
            logger.info(f"Making {method} request to {endpoint}")
            if method == 'GET':
                response = requests.get(url, headers=self.headers, params=params, timeout=60)
            elif method == 'POST':
                response = requests.post(url, headers=self.headers, json=data, timeout=60)
            else:
                raise ValueError(f"Unsupported method: {method}")

            # Log the xtm-trace-id header for support purposes
            trace_id = response.headers.get('xtm-trace-id', 'N/A')
            logger.info(f"XTM Trace ID: {trace_id} | Endpoint: {endpoint} | Status: {response.status_code}")

            response.raise_for_status()
            return response.json() if response.content else {}
        except requests.exceptions.RequestException as e:
            # Also log trace ID on error if available
            trace_id = 'N/A'
            if hasattr(e, 'response') and e.response is not None:
                trace_id = e.response.headers.get('xtm-trace-id', 'N/A')
                logger.error(f"API request failed | XTM Trace ID: {trace_id} | Error: {e}")
            else:
                logger.error(f"API request failed: {e}")
            raise

    def get_projects(self, status: str = None) -> List[Dict]:
        """Retrieve projects from XTM API."""
        try:
            # Get projects - adjust endpoint based on actual XTM API
            params = {}
            if status:
                params['status'] = status

            projects = self._make_request('projects', params=params)
            logger.info(f"Retrieved {len(projects) if isinstance(projects, list) else 'unknown'} projects")
            return projects if isinstance(projects, list) else []
        except Exception as e:
            logger.error(f"Failed to get projects: {e}")
            return []

    def get_project_metrics(self, project_id: int):
        """Get detailed metrics for a specific project (returns list)."""
        try:
            metrics = self._make_request(f'projects/{project_id}/metrics')
            # Metrics is a list of target languages
            if isinstance(metrics, list):
                return metrics
            elif isinstance(metrics, dict):
                return [metrics]  # Wrap single dict in list for consistency
            return []
        except Exception as e:
            logger.warning(f"Failed to get metrics for project {project_id}: {e}")
            return []

    def get_project_statistics(self, project_id: int, excluded_users: List[str] = None):
        """Get detailed per-user statistics for a project, excluding specified users."""
        if excluded_users is None:
            excluded_users = ["leo.chang@familysearch.org", "LeoAdmin", "Robert.Sena@churchofjesuschrist.org", "MartinADMIN"]

        try:
            stats = self._make_request(f'projects/{project_id}/statistics')
            if not isinstance(stats, list):
                return []

            # Convert excluded users to lowercase for case-insensitive comparison
            excluded_users_lower = [user.lower() for user in excluded_users]

            # Filter out excluded users' work
            filtered_stats = []
            for lang_stats in stats:
                users_statistics = lang_stats.get('usersStatistics', [])
                # Keep only users who are NOT in the excluded list
                filtered_users = [
                    user for user in users_statistics
                    if user.get('username', '').lower() not in excluded_users_lower
                ]

                if filtered_users:
                    lang_stats_copy = lang_stats.copy()
                    lang_stats_copy['usersStatistics'] = filtered_users
                    filtered_stats.append(lang_stats_copy)

            return filtered_stats
        except Exception as e:
            logger.warning(f"Failed to get statistics for project {project_id}: {e}")
            return []

    def get_workflow_steps(self, project_id: int) -> List[Dict]:
        """Get workflow steps for a project."""
        try:
            # Try different endpoints for workflow data
            try:
                steps = self._make_request(f'projects/{project_id}/jobs')
                return steps if isinstance(steps, list) else []
            except:
                # Fallback to user workflow steps
                steps = self._make_request(f'projects/{project_id}/workflow')
                return steps if isinstance(steps, list) else []
        except Exception as e:
            logger.debug(f"No workflow data available for project {project_id}: {e}")
            return []

    def get_users_workflow_steps(self) -> List[Dict]:
        """Get user workflow step assignments."""
        try:
            steps = self._make_request('users/workflow-steps')
            return steps if isinstance(steps, list) else []
        except Exception as e:
            logger.warning(f"Failed to get user workflow steps: {e}")
            return []

    def aggregate_monthly_data(self, start_month: str = None, end_month: str = None) -> Dict:
        """Aggregate data for the specified date range."""
        if start_month is None:
            start_month = self.report_month
        if end_month is None:
            end_month = self.report_month

        logger.info(f"Aggregating data from {start_month} to {end_month}")

        # Initialize data structures
        data = {
            'project_stats': {
                'total': 0,
                'completed': 0,
                'in_progress': 0,
                'pending': 0
            },
            'workflow_by_language': {},  # New: workflow metrics per language
            'user_statistics': {},  # User-level statistics
            'projects': []
        }

        # Get all projects
        projects = self.get_projects()

        for project in projects:
            project_id = project.get('id')
            if not project_id:
                continue

            # Get project statistics (filtered to exclude specified users)
            # Wrap in try-except to continue even if individual projects fail
            try:
                stats_list = self.get_project_statistics(project_id)
            except Exception as e:
                logger.warning(f"Failed to get statistics for project {project_id}, skipping: {e}")
                continue

            # Process each target language in the statistics
            if isinstance(stats_list, list) and stats_list:
                project_total_words = 0
                project_has_work_in_period = False
                target_languages = []
                source_lang = 'en_US'  # Default, will try to determine from context

                for lang_stats in stats_list:
                    target_lang_code = lang_stats.get('targetLanguage', 'unknown')
                    target_lang_name = self._locale_to_language_name(target_lang_code)

                    # Aggregate words from all users (already filtered to exclude specified users)
                    users_statistics = lang_stats.get('usersStatistics', [])

                    for user_stats in users_statistics:
                        username = user_stats.get('username', 'Unknown')
                        steps_statistics = user_stats.get('stepsStatistics', [])

                        for step_stat in steps_statistics:
                            step_name = step_stat.get('workflowStepName', '')
                            # Remove numbers from step name (e.g., "translate1" -> "translate")
                            clean_step_name = ''.join([c for c in step_name if not c.isdigit()])

                            # Get total words from all jobs for this user/step/language
                            jobs_statistics = step_stat.get('jobsStatistics', [])
                            step_total_words = 0

                            for job_stat in jobs_statistics:
                                # Filter by finish date (lastCompletionDate)
                                # If blank, ignore this job
                                last_completion_date = job_stat.get('lastCompletionDate')
                                completion_month = None

                                if last_completion_date:
                                    try:
                                        if isinstance(last_completion_date, (int, float)):
                                            dt = datetime.fromtimestamp(last_completion_date / 1000)
                                            completion_month = dt.strftime('%Y-%m')
                                    except:
                                        pass

                                # Skip work not completed in the target date range
                                if completion_month:
                                    if completion_month < start_month or completion_month > end_month:
                                        continue
                                else:
                                    # If no completion date, skip this job
                                    continue

                                # Calculate total words using all match type fields:
                                # iceMatchWords + leveragedWords + repeatsWords + machineTranslationWords +
                                # highFuzzyMatchWords + mediumFuzzyMatchWords + lowFuzzyMatchWords +
                                # highFuzzyRepeatsWords + mediumFuzzyRepeatsWords + lowFuzzyRepeatsWords +
                                # noMatchingWords
                                source_stats = job_stat.get('sourceStatistics', {})
                                ice_match_words = source_stats.get('iceMatchWords', 0)
                                leveraged_words = source_stats.get('leveragedWords', 0)
                                repeats_words = source_stats.get('repeatsWords', 0)
                                machine_translation_words = source_stats.get('machineTranslationWords', 0)
                                high_fuzzy_match_words = source_stats.get('highFuzzyMatchWords', 0)
                                medium_fuzzy_match_words = source_stats.get('mediumFuzzyMatchWords', 0)
                                low_fuzzy_match_words = source_stats.get('lowFuzzyMatchWords', 0)
                                high_fuzzy_repeats_words = source_stats.get('highFuzzyRepeatsWords', 0)
                                medium_fuzzy_repeats_words = source_stats.get('mediumFuzzyRepeatsWords', 0)
                                low_fuzzy_repeats_words = source_stats.get('lowFuzzyRepeatsWords', 0)
                                no_matching_words = source_stats.get('noMatchingWords', 0)

                                job_words = (ice_match_words + leveraged_words + repeats_words + machine_translation_words +
                                           high_fuzzy_match_words + medium_fuzzy_match_words + low_fuzzy_match_words +
                                           high_fuzzy_repeats_words + medium_fuzzy_repeats_words + low_fuzzy_repeats_words +
                                           no_matching_words)
                                step_total_words += job_words

                            # Mark that this project has work
                            if step_total_words > 0:
                                project_has_work_in_period = True

                            # Only aggregate if this step had words in the target period
                            if step_total_words > 0:
                                # Add language to list if not already there
                                if target_lang_name not in target_languages:
                                    target_languages.append(target_lang_name)

                                # Create unique key for workflow step + language (using language name)
                                workflow_key = f"{clean_step_name} - {target_lang_name}"

                                if workflow_key not in data['workflow_by_language']:
                                    data['workflow_by_language'][workflow_key] = {
                                        'workflow_step': clean_step_name,
                                        'language': target_lang_name,
                                        'words_done': 0,
                                        'words_to_be_done': 0,
                                        'projects': 0
                                    }

                                data['workflow_by_language'][workflow_key]['words_done'] += step_total_words
                                project_total_words += step_total_words

                                # Aggregate user statistics
                                user_lang_key = f"{username}||{target_lang_name}"
                                if user_lang_key not in data['user_statistics']:
                                    data['user_statistics'][user_lang_key] = {
                                        'username': username,
                                        'language': target_lang_name,
                                        'workflow_steps': {}
                                    }

                                if clean_step_name not in data['user_statistics'][user_lang_key]['workflow_steps']:
                                    data['user_statistics'][user_lang_key]['workflow_steps'][clean_step_name] = 0

                                data['user_statistics'][user_lang_key]['workflow_steps'][clean_step_name] += step_total_words

                # Only add project to stats if it has work in the target period
                if project_has_work_in_period and project_total_words > 0:
                    # Update project stats
                    data['project_stats']['total'] += 1
                    status = project.get('status', 'UNKNOWN')

                    if status == 'FINISHED':
                        data['project_stats']['completed'] += 1
                    elif status in ['IN_PROGRESS', 'STARTED']:
                        data['project_stats']['in_progress'] += 1
                    else:
                        data['project_stats']['pending'] += 1

                    # Mark projects for each language/workflow combination
                    for target_lang_name in target_languages:
                        for step_name in set([s.get('workflowStepName', '') for lang_stat in stats_list for u in lang_stat.get('usersStatistics', []) for s in u.get('stepsStatistics', [])]):
                            clean_step_name = ''.join([c for c in step_name if not c.isdigit()])
                            workflow_key = f"{clean_step_name} - {target_lang_name}"
                            if workflow_key in data['workflow_by_language']:
                                data['workflow_by_language'][workflow_key]['projects'] = data['workflow_by_language'][workflow_key].get('projects', 0) + 1

                    # Store project summary
                    data['projects'].append({
                        'id': project_id,
                        'name': project.get('name', 'Unknown'),
                        'status': status,
                        'source_lang': source_lang,
                        'target_langs': ', '.join(target_languages),
                        'total_words': project_total_words,
                        'created_date': None  # Filtering by finish date, not project creation
                    })

        logger.info(f"Aggregated data for {data['project_stats']['total']} projects")
        return data

    def create_workflow_sheet(self, wb, sheet_name: str, data: Dict, title: str):
        """Create a workflow sheet with data and chart."""
        if sheet_name == "Monthly":
            ws_workflow = wb.active
            ws_workflow.title = sheet_name
        else:
            ws_workflow = wb.create_sheet(sheet_name)

        # Define styles
        header_font = Font(bold=True, color="FFFFFF", size=12)
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        ws_workflow['A1'] = title
        ws_workflow['A1'].font = Font(bold=True, size=14)

        # First, organize data by language and workflow step
        language_workflow_data = {}
        workflow_steps = set()

        for workflow_key, metrics in data['workflow_by_language'].items():
            language = metrics['language']
            workflow_step = metrics['workflow_step']
            words_done = metrics['words_done']

            workflow_steps.add(workflow_step)

            if language not in language_workflow_data:
                language_workflow_data[language] = {}

            language_workflow_data[language][workflow_step] = words_done

        # Define specific column order for workflow steps
        workflow_order = ['translate', 'correct', 'final review']
        # Only include workflow steps that exist in the data, in the specified order
        sorted_workflow_steps = [step for step in workflow_order if step in workflow_steps]
        # Add any workflow steps not in the predefined order at the end (sorted alphabetically)
        remaining_steps = sorted([step for step in workflow_steps if step not in workflow_order])
        sorted_workflow_steps.extend(remaining_steps)

        # Create header row: Language | Workflow Step 1 | Workflow Step 2 | ... | Total
        headers = ['Language'] + sorted_workflow_steps + ['Total']
        for col_idx, header in enumerate(headers, start=1):
            cell = ws_workflow.cell(row=3, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = border

        # Fill in data rows
        row_idx = 4
        for language in sorted(language_workflow_data.keys()):
            row_data = [language]
            row_total = 0

            # Add words done for each workflow step
            for workflow_step in sorted_workflow_steps:
                words_done = language_workflow_data[language].get(workflow_step, 0)
                row_data.append(words_done)
                row_total += words_done

            # Add total column
            row_data.append(row_total)

            for col_idx, value in enumerate(row_data, start=1):
                cell = ws_workflow.cell(row=row_idx, column=col_idx, value=value)
                cell.border = border
                # Bold the total column
                if col_idx == len(row_data):
                    cell.font = Font(bold=True)

            row_idx += 1

        # Adjust column widths
        ws_workflow.column_dimensions['A'].width = 15
        for col_idx in range(2, len(headers) + 1):
            ws_workflow.column_dimensions[get_column_letter(col_idx)].width = 18

        # Enable AutoFilter for sorting/filtering
        last_col_letter = get_column_letter(len(headers))
        last_data_row = row_idx - 1
        ws_workflow.auto_filter.ref = f"A3:{last_col_letter}{last_data_row}"

        # Add bar chart showing only totals
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Total Words Processed by Language"
        chart.y_axis.title = "Total Words"
        chart.x_axis.title = "Language"

        num_languages = len(language_workflow_data)

        # Categories (languages) - column A, starting from row 4
        cats = Reference(ws_workflow, min_col=1, min_row=4, max_row=3 + num_languages)

        # Data series - only the Total column (last column)
        total_col_idx = len(headers)  # Total is the last column
        data = Reference(ws_workflow, min_col=total_col_idx, min_row=3, max_row=3 + num_languages)
        chart.add_data(data, titles_from_data=True)

        chart.set_categories(cats)
        chart.height = 15
        chart.width = 25

        # Position chart to the right of the table
        ws_workflow.add_chart(chart, "G3")

    def create_user_statistics_sheet(self, wb, sheet_name: str, data: Dict, title: str):
        """Create a user statistics sheet with user names, languages, and workflow steps."""
        ws_users = wb.create_sheet(sheet_name)

        # Define styles
        header_font = Font(bold=True, color="FFFFFF", size=12)
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        ws_users['A1'] = title
        ws_users['A1'].font = Font(bold=True, size=14)

        # Collect all workflow steps from the data
        workflow_steps = set()
        for user_key, user_data in data['user_statistics'].items():
            workflow_steps.update(user_data['workflow_steps'].keys())

        # Define specific column order for workflow steps
        workflow_order = ['translate', 'correct', 'final review']
        # Only include workflow steps that exist in the data, in the specified order
        sorted_workflow_steps = [step for step in workflow_order if step in workflow_steps]
        # Add any workflow steps not in the predefined order at the end (sorted alphabetically)
        remaining_steps = sorted([step for step in workflow_steps if step not in workflow_order])
        sorted_workflow_steps.extend(remaining_steps)

        # Create header row: Name | Language | Workflow Step 1 | Workflow Step 2 | ... | Total
        headers = ['Name', 'Language'] + sorted_workflow_steps + ['Total']
        for col_idx, header in enumerate(headers, start=1):
            cell = ws_users.cell(row=3, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = border

        # Sort users by total words (descending)
        user_list = []
        for user_key, user_data in data['user_statistics'].items():
            total_words = sum(user_data['workflow_steps'].values())
            user_list.append((user_data['username'], user_data['language'], user_data['workflow_steps'], total_words))

        user_list.sort(key=lambda x: x[3], reverse=True)

        # Fill in data rows
        row_idx = 4
        for username, language, workflow_steps_data, total_words in user_list:
            row_data = [username, language]

            # Add words done for each workflow step
            for workflow_step in sorted_workflow_steps:
                words_done = workflow_steps_data.get(workflow_step, 0)
                row_data.append(words_done)

            # Add total column
            row_data.append(total_words)

            for col_idx, value in enumerate(row_data, start=1):
                cell = ws_users.cell(row=row_idx, column=col_idx, value=value)
                cell.border = border
                # Bold the total column
                if col_idx == len(row_data):
                    cell.font = Font(bold=True)

            row_idx += 1

        # Adjust column widths
        ws_users.column_dimensions['A'].width = 40
        ws_users.column_dimensions['B'].width = 25
        for col_idx in range(3, len(headers) + 1):
            ws_users.column_dimensions[get_column_letter(col_idx)].width = 15

        # Enable AutoFilter for sorting/filtering
        last_col_letter = get_column_letter(len(headers))
        last_data_row = row_idx - 1
        ws_users.auto_filter.ref = f"A3:{last_col_letter}{last_data_row}"

    def create_excel_report(self, monthly_data: Dict, ytd_data: Dict, output_path: str) -> str:
        """Create Excel report with monthly and YTD sheets."""
        logger.info("Creating Excel report")

        wb = Workbook()

        # Create Monthly sheet
        self.create_workflow_sheet(wb, "Monthly", monthly_data, f"Monthly Report - {self.report_month_name}")

        # Create Monthly User Statistics sheet
        if monthly_data.get('user_statistics'):
            self.create_user_statistics_sheet(wb, "User Statistics - Monthly", monthly_data, f"User Statistics - {self.report_month_name}")

        # Create YTD sheet
        self.create_workflow_sheet(wb, "Year-to-Date", ytd_data, f"Year-to-Date Report - {self.ytd_start_month} to {self.ytd_end_month}")

        # Create YTD User Statistics sheet
        if ytd_data.get('user_statistics'):
            self.create_user_statistics_sheet(wb, "User Statistics - YTD", ytd_data, f"User Statistics - YTD ({self.ytd_start_month} to {self.ytd_end_month})")

        # Save workbook with fallback locations
        save_successful = False
        last_error = None

        # Try primary location
        try:
            wb.save(output_path)
            logger.info(f"Excel report saved to {output_path}")
            save_successful = True
            return output_path
        except Exception as e:
            last_error = e
            logger.warning(f"Failed to save to primary location {output_path}: {e}")

        # Try fallback location 1: Desktop
        if not save_successful:
            try:
                fallback_path = Path.home() / "Desktop" / Path(output_path).name
                wb.save(str(fallback_path))
                logger.info(f"Excel report saved to fallback location: {fallback_path}")
                save_successful = True
                return str(fallback_path)
            except Exception as e:
                last_error = e
                logger.warning(f"Failed to save to Desktop fallback: {e}")

        # Try fallback location 2: Current working directory
        if not save_successful:
            try:
                fallback_path = Path.cwd() / Path(output_path).name
                wb.save(str(fallback_path))
                logger.info(f"Excel report saved to fallback location: {fallback_path}")
                save_successful = True
                return str(fallback_path)
            except Exception as e:
                last_error = e
                logger.warning(f"Failed to save to current directory fallback: {e}")

        # Try fallback location 3: Temp directory
        if not save_successful:
            try:
                import tempfile
                fallback_path = Path(tempfile.gettempdir()) / Path(output_path).name
                wb.save(str(fallback_path))
                logger.info(f"Excel report saved to temp directory fallback: {fallback_path}")
                save_successful = True
                return str(fallback_path)
            except Exception as e:
                last_error = e
                logger.error(f"Failed to save to temp directory fallback: {e}")

        # If all save attempts failed, raise the last error
        if not save_successful:
            raise Exception(f"Failed to save Excel report to any location. Last error: {last_error}")

    def _calculate_summary_stats(self, data: Dict) -> Dict:
        """Calculate summary statistics from the data."""
        stats = {
            'total_words': 0,
            'top_languages': '',
            'workflow_summary': ''
        }

        # Calculate total words and get top languages
        language_totals = {}
        for workflow_key, metrics in data['workflow_by_language'].items():
            language = metrics['language']
            words_done = metrics['words_done']

            if language not in language_totals:
                language_totals[language] = 0
            language_totals[language] += words_done
            stats['total_words'] += words_done

        # Get top 3 languages
        top_languages = sorted(language_totals.items(), key=lambda x: x[1], reverse=True)[:3]
        top_langs_text = []
        for i, (lang, words) in enumerate(top_languages, 1):
            top_langs_text.append(f"  {i}. {lang}: {words:,} words")
        stats['top_languages'] = '\n'.join(top_langs_text) if top_langs_text else "  No data available"

        # Calculate workflow breakdown
        workflow_totals = {}
        for workflow_key, metrics in data['workflow_by_language'].items():
            workflow_step = metrics['workflow_step']
            words_done = metrics['words_done']

            if workflow_step not in workflow_totals:
                workflow_totals[workflow_step] = 0
            workflow_totals[workflow_step] += words_done

        workflow_text = []
        for workflow_step, words in sorted(workflow_totals.items()):
            workflow_text.append(f"- {workflow_step}: {words:,} words")
        stats['workflow_summary'] = '\n'.join(workflow_text) if workflow_text else "No data available"

        return stats

    def _send_system_notification(self, title: str, message: str, sound: bool = True):
        """Send a macOS system notification."""
        try:
            script = f'''
            display notification "{message}" with title "{title}"'''
            if sound:
                script += ' sound name "Glass"'

            subprocess.run(
                ['osascript', '-e', script],
                capture_output=True,
                text=True,
                timeout=5
            )
            logger.info(f"System notification sent: {title}")
        except Exception as e:
            logger.warning(f"Could not send system notification: {e}")

    def _send_failure_notification(self, error_message: str, report_path: str = None):
        """Send notification about failure via multiple channels."""
        logger.info("Sending failure notifications")

        # 1. Send macOS system notification
        self._send_system_notification(
            "XTM Report Failed",
            f"Report generation encountered an error. Check logs for details.",
            sound=True
        )

        # 2. Try to send email notification
        try:
            subject = f"ALERT: XTM Monthly Report Failed - {self.report_month}"
            body = f"""ALERT: The automated XTM monthly report generation has failed.

Report Period: {self.report_month_name}
Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Error: {error_message}

Please check the log file at:
{Path.cwd() / 'xtm_report.log'}
"""
            if report_path:
                body += f"\nPartial report may be available at: {report_path}"

            body += "\n\nPlease investigate and re-run manually if needed."

            # Escape for AppleScript
            subject_escaped = subject.replace('\\', '\\\\').replace('"', '\\"')
            body_escaped = body.replace('\\', '\\\\').replace('"', '\\"')

            recipients = self.config.get('email_recipients', [])
            if recipients:
                # Try simple notification email via Mail app
                recipients_script = ""
                for recipient in recipients:
                    recipient_escaped = recipient.replace('\\', '\\\\').replace('"', '\\"')
                    recipients_script += f'make new to recipient at end of to recipients of new_message with properties {{address:"{recipient_escaped}"}}\n'

                applescript = f'''
                tell application "Mail"
                    set new_message to make new outgoing message with properties {{subject:"{subject_escaped}", visible:false}}
                    tell new_message
                        {recipients_script}
                        set the content to "{body_escaped}"
                    end tell
                    send new_message
                end tell
                '''

                subprocess.run(
                    ['osascript', '-e', applescript],
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                logger.info("Failure notification email sent")
        except Exception as e:
            logger.warning(f"Could not send failure email notification: {e}")

    def _ensure_outlook_running(self) -> bool:
        """Ensure Microsoft Outlook is running, launch if needed."""
        try:
            # Check if Outlook is running
            check_script = '''
            tell application "System Events"
                return (name of processes) contains "Microsoft Outlook"
            end tell
            '''

            result = subprocess.run(
                ['osascript', '-e', check_script],
                capture_output=True,
                text=True,
                timeout=10
            )

            is_running = result.stdout.strip() == "true"

            if is_running:
                logger.info("Microsoft Outlook is already running")
                return True

            # Launch Outlook
            logger.info("Microsoft Outlook is not running. Launching...")
            launch_script = '''
            tell application "Microsoft Outlook"
                activate
            end tell
            '''

            result = subprocess.run(
                ['osascript', '-e', launch_script],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode == 0:
                # Wait for Outlook to fully start (up to 30 seconds)
                logger.info("Waiting for Outlook to start...")
                import time
                for i in range(30):
                    time.sleep(1)
                    check_result = subprocess.run(
                        ['osascript', '-e', check_script],
                        capture_output=True,
                        text=True,
                        timeout=10
                    )
                    if check_result.stdout.strip() == "true":
                        logger.info(f"Outlook started successfully after {i+1} seconds")
                        # Give it a couple more seconds to fully initialize
                        time.sleep(3)
                        return True

                logger.warning("Outlook did not start within 30 seconds")
                return False
            else:
                logger.warning(f"Failed to launch Outlook: {result.stderr}")
                return False

        except Exception as e:
            logger.warning(f"Error checking/launching Outlook: {e}")
            return False

    def send_email_via_outlook(self, report_path: str, monthly_data: Dict, ytd_data: Dict):
        """Create and open email draft in Outlook (macOS version)."""
        try:
            logger.info("Preparing to send email via Outlook")

            # Ensure Outlook is running (especially important for auto-send)
            if self.auto_send:
                if not self._ensure_outlook_running():
                    logger.error("Could not ensure Outlook is running. Falling back to Apple Mail.")
                    # Fall through to Apple Mail attempt below

            # Calculate summary statistics
            monthly_stats = self._calculate_summary_stats(monthly_data)
            ytd_stats = self._calculate_summary_stats(ytd_data)

            # Prepare email content
            subject = f"XTM Monthly Report - {self.report_month}"

            body = f"""Hello,

Please find attached the XTM monthly report for {self.report_month}.

Monthly Summary ({self.report_month_name}):
- Total Words Processed: {monthly_stats['total_words']:,}
- Top Languages:
{monthly_stats['top_languages']}

Year-to-Date Summary ({self.ytd_start_month} to {self.ytd_end_month}):
- Total Words Processed: {ytd_stats['total_words']:,}
- Top Languages:
{ytd_stats['top_languages']}

Report Generated: {self.report_date.strftime('%Y-%m-%d %H:%M')}

Please review and let me know if you have any questions.

Best regards
"""

            # Escape quotes and backslashes for AppleScript
            subject_escaped = subject.replace('\\', '\\\\').replace('"', '\\"')
            body_escaped = body.replace('\\', '\\\\').replace('"', '\\"')
            report_path_posix = Path(report_path).as_posix()

            recipients = self.config['email_recipients']

            # Try Microsoft Outlook first (already ensured it's running if auto_send)
            outlook_result = self._create_outlook_email_mac(subject_escaped, body_escaped, recipients, report_path_posix)
            if outlook_result:
                return

            # Fall back to Apple Mail
            logger.info("Microsoft Outlook not available, trying Apple Mail")
            apple_result = self._create_apple_mail_email(subject_escaped, body_escaped, recipients, report_path_posix)
            if apple_result:
                return

            # If both fail, just open the report
            logger.warning("Could not create email draft. Opening report file.")
            subprocess.run(['open', report_path])
            print(f"\n⚠ Email draft creation failed. Report opened: {report_path}")
            print(f"Recipients: {', '.join(recipients)}")

        except Exception as e:
            logger.error(f"Failed to create email: {e}")
            logger.info(f"Report saved to: {report_path}")
            raise

    def _create_outlook_email_mac(self, subject: str, body: str, recipients: List[str], attachment_path: str) -> bool:
        """Create email draft in Microsoft Outlook for Mac using AppleScript."""
        try:
            # Build recipient list for AppleScript - simpler syntax
            recipients_script = ""
            for recipient in recipients:
                recipient_escaped = recipient.replace('\\', '\\\\').replace('"', '\\"')
                recipients_script += f'make new to recipient with properties {{email address:{{address:"{recipient_escaped}"}}}}\n'

            # Different AppleScript depending on whether we're sending or just creating draft
            if self.auto_send:
                applescript = f'''
                tell application "Microsoft Outlook"
                    set new_message to make new outgoing message with properties {{subject:"{subject}", content:"{body}"}}
                    tell new_message
                        {recipients_script}
                        make new attachment with properties {{file:POSIX file "{attachment_path}"}}
                        send
                    end tell
                end tell
                '''
            else:
                applescript = f'''
                tell application "Microsoft Outlook"
                    set new_message to make new outgoing message with properties {{subject:"{subject}", content:"{body}"}}
                    tell new_message
                        {recipients_script}
                        make new attachment with properties {{file:POSIX file "{attachment_path}"}}
                    end tell
                    open new_message
                    activate
                end tell
                '''

            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode == 0:
                if self.auto_send:
                    logger.info("Microsoft Outlook email sent successfully")
                else:
                    logger.info("Microsoft Outlook email draft created successfully")
                return True
            else:
                logger.warning(f"Outlook AppleScript failed: {result.stderr}")
                return False

        except Exception as e:
            logger.warning(f"Failed to create Outlook email on Mac: {e}")
            return False

    def _create_apple_mail_email(self, subject: str, body: str, recipients: List[str], attachment_path: str) -> bool:
        """Create email draft in Apple Mail using AppleScript."""
        try:
            # Build recipient list for AppleScript
            recipients_script = ""
            for recipient in recipients:
                recipient_escaped = recipient.replace('\\', '\\\\').replace('"', '\\"')
                recipients_script += f'make new to recipient at end of to recipients of new_message with properties {{address:"{recipient_escaped}"}}\n'

            # Different AppleScript depending on whether we're sending or just creating draft
            # Use simple attachment without positioning to avoid HTML file creation
            if self.auto_send:
                applescript = f'''
                tell application "Mail"
                    set new_message to make new outgoing message with properties {{subject:"{subject}", sender:"leo.chang@familysearch.org"}}
                    tell new_message
                        {recipients_script}
                        set the content to "{body}"
                        make new attachment with properties {{file name:POSIX file "{attachment_path}"}}
                    end tell
                    send new_message
                end tell
                '''
            else:
                applescript = f'''
                tell application "Mail"
                    set new_message to make new outgoing message with properties {{subject:"{subject}", visible:true, sender:"leo.chang@familysearch.org"}}
                    tell new_message
                        {recipients_script}
                        set the content to "{body}"
                        make new attachment with properties {{file name:POSIX file "{attachment_path}"}}
                    end tell
                    activate
                end tell
                '''

            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode == 0:
                if self.auto_send:
                    logger.info("Apple Mail email sent successfully")
                else:
                    logger.info("Apple Mail email draft created successfully")
                return True
            else:
                logger.warning(f"Apple Mail AppleScript failed: {result.stderr}")
                return False

        except Exception as e:
            logger.warning(f"Failed to create Apple Mail email: {e}")
            return False

    def generate_report(self):
        """Main method to generate and distribute the report."""
        report_path = None
        try:
            logger.info("Starting XTM monthly report generation")

            # Run health checks
            self._run_health_checks()

            # Aggregate monthly data
            try:
                monthly_data = self.aggregate_monthly_data(self.report_month, self.report_month)
            except Exception as e:
                logger.error(f"Failed to aggregate monthly data: {e}", exc_info=True)
                # Initialize with empty data structure
                monthly_data = {
                    'project_stats': {'total': 0, 'completed': 0, 'in_progress': 0, 'pending': 0},
                    'workflow_by_language': {},
                    'user_statistics': {},
                    'projects': []
                }

            # Aggregate year-to-date data
            try:
                ytd_data = self.aggregate_monthly_data(self.ytd_start_month, self.ytd_end_month)
            except Exception as e:
                logger.error(f"Failed to aggregate YTD data: {e}", exc_info=True)
                # Use monthly data as fallback
                ytd_data = monthly_data

            # Validate we have at least some data
            has_data = (monthly_data.get('workflow_by_language') or
                       ytd_data.get('workflow_by_language'))

            if not has_data:
                logger.warning("No data available for reporting period")

            # Create output filename
            output_filename = f"XTM_Monthly_Report_{self.report_month}_{self.report_date.strftime('%Y%m%d')}.xlsx"
            output_path = Path(self.config['onedrive_path']) / output_filename

            # Ensure output directory exists (but don't fail if we can't create it)
            try:
                output_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                logger.warning(f"Could not create output directory: {e}")

            # Create Excel report
            try:
                report_path = self.create_excel_report(monthly_data, ytd_data, str(output_path))
            except Exception as e:
                logger.error(f"Failed to create Excel report: {e}", exc_info=True)
                raise

            # Send via Outlook (pass the data for email body)
            # Don't fail the entire process if email fails
            try:
                self.send_email_via_outlook(report_path, monthly_data, ytd_data)
                email_success = True
            except Exception as e:
                logger.error(f"Failed to send email: {e}", exc_info=True)
                email_success = False

            logger.info("Report generation completed successfully")
            print(f"\n✓ Report generated: {report_path}")
            print(f"✓ Monthly period: {self.report_month_name}")
            print(f"✓ YTD period: {self.ytd_start_month} to {self.ytd_end_month}")
            if email_success:
                if self.auto_send:
                    print(f"✓ Email sent automatically to {len(self.config['email_recipients'])} recipients")
                else:
                    print(f"✓ Email draft created with {len(self.config['email_recipients'])} recipients")
            else:
                print(f"⚠ Email could not be sent - report saved locally")

        except Exception as e:
            logger.error(f"Report generation failed: {e}", exc_info=True)
            # Send failure notification
            try:
                self._send_failure_notification(str(e), report_path)
            except:
                pass
            raise


def main():
    """Main entry point."""
    import argparse

    parser = argparse.ArgumentParser(description='Generate XTM monthly report')
    parser.add_argument('--auto-send', action='store_true',
                       help='Automatically send email (default: create draft only)')
    args = parser.parse_args()

    try:
        generator = XTMReportGenerator(auto_send=args.auto_send)
        generator.generate_report()
    except Exception as e:
        logger.error(f"Fatal error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
