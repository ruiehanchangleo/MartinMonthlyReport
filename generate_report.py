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
        'en_US': 'English (USA)',
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
        'fr_FR': 'French (France)',
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
        'swa': 'Swahili',
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

    # Users excluded from all reports (checked case-insensitively against all name fields)
    EXCLUDED_USERS = [
        "leo.chang@familysearch.org", "LeoAdmin",
        "Robert.Sena@churchofjesuschrist.org", "MartinADMIN",
        "Tester BSP BSP", "BSP BSP Tester", "BSP_Tester",
        "ben.randall@brightspot.com"
    ]

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
        """Convert locale code to base language name, combining variants."""
        name = self.LOCALE_TO_LANGUAGE.get(locale_code, locale_code)
        # Strip parenthetical variants: "Portuguese (Brazil)" -> "Portuguese"
        if '(' in name:
            name = name.split('(')[0].strip()
        return name

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

    def get_project_metrics_data(self, project_id: int):
        """
        Get project metrics showing wordsDone for each workflow step.
        Returns metrics for all languages in the project with wordsDone by workflow step.
        This shows CURRENT progress, not historical monthly completions.
        """
        try:
            metrics = self._make_request(f'projects/{project_id}/metrics')
            if not isinstance(metrics, list):
                return []
            return metrics
        except Exception as e:
            logger.warning(f"Failed to get metrics for project {project_id}: {e}")
            return []

    def get_project_status_with_steps(self, project_id: int):
        """
        Get project status with workflow step completion dates.
        This endpoint preserves finishDate even when work is reopened.
        Returns job-level data with step-level finish dates.
        """
        try:
            status = self._make_request(f'projects/{project_id}/status?fetchLevel=STEPS')
            if not isinstance(status, dict):
                return {}
            return status
        except Exception as e:
            logger.warning(f"Failed to get status for project {project_id}: {e}")
            return {}

    def _is_excluded_user(self, user_stat: Dict) -> bool:
        """Check if a user should be excluded, matching against all name fields."""
        excluded_lower = [u.lower() for u in self.EXCLUDED_USERS]

        # Check all possible name fields
        fields_to_check = [
            user_stat.get('username', ''),
            user_stat.get('userDisplayName', ''),
        ]

        # Also check firstName + lastName combination
        first_name = user_stat.get('firstName', '').strip()
        last_name = user_stat.get('lastName', '').strip()
        if first_name and last_name:
            fields_to_check.append(f"{first_name} {last_name}")
        if first_name:
            fields_to_check.append(first_name)
        if last_name:
            fields_to_check.append(last_name)

        return any(field.lower() in excluded_lower for field in fields_to_check if field)

    @staticmethod
    def _resolve_user_name(user_stat: Dict) -> str:
        """Extract the best display name from a user statistics entry."""
        first_name = user_stat.get('firstName', '').strip()
        last_name = user_stat.get('lastName', '').strip()

        if first_name and last_name:
            return f"{first_name} {last_name}"
        if first_name:
            return first_name
        if last_name:
            return last_name

        # Fall back to display name or email
        username = user_stat.get('userDisplayName', user_stat.get('username', 'Unknown'))
        # Strip "generic " prefix from XTM display names
        if username.lower().startswith('generic '):
            username = username[8:]
        return username

    def get_project_statistics(self, project_id: int):
        """Get detailed per-user statistics for a project, excluding specified users."""
        try:
            stats = self._make_request(f'projects/{project_id}/statistics')
            if not isinstance(stats, list):
                return []

            # Filter out excluded users' work
            filtered_stats = []
            for lang_stats in stats:
                users_statistics = lang_stats.get('usersStatistics', [])
                filtered_users = [
                    user for user in users_statistics
                    if not self._is_excluded_user(user)
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
        """Aggregate data for the specified date range using completion dates."""
        if start_month is None:
            start_month = self.report_month
        if end_month is None:
            end_month = self.report_month

        logger.info(f"Aggregating data from {start_month} to {end_month}")

        # Parse date range
        from datetime import datetime
        start_date = datetime.strptime(start_month + '-01', '%Y-%m-%d')
        # End date is last day of end_month
        if end_month != start_month:
            end_year, end_month_num = map(int, end_month.split('-'))
            if end_month_num == 12:
                end_date = datetime(end_year + 1, 1, 1)
            else:
                end_date = datetime(end_year, end_month_num + 1, 1)
        else:
            year, month = map(int, start_month.split('-'))
            if month == 12:
                end_date = datetime(year + 1, 1, 1)
            else:
                end_date = datetime(year, month + 1, 1)

        logger.info(f"Date range: {start_date} to {end_date}")

        # Initialize data structures
        data = {
            'project_stats': {
                'total': 0,
                'completed': 0,
                'in_progress': 0,
                'pending': 0
            },
            'workflow_by_language': {},  # workflow metrics per language
            'user_statistics': {},  # User-level statistics
            'projects': []
        }

        # Get all projects
        projects = self.get_projects()

        # Fetch statistics and status for all projects in parallel
        from concurrent.futures import ThreadPoolExecutor, as_completed

        def fetch_project_data(project_id):
            try:
                stats = self.get_project_statistics(project_id)
                status = self.get_project_status_with_steps(project_id)
                return project_id, stats, status
            except Exception as e:
                logger.warning(f"Failed to get data for project {project_id}: {e}")
                return project_id, [], {}

        project_ids = [p.get('id') for p in projects if p.get('id')]
        project_map = {p.get('id'): p for p in projects if p.get('id')}
        project_results = {}

        logger.info(f"Fetching data for {len(project_ids)} projects using parallel requests...")
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {executor.submit(fetch_project_data, pid): pid for pid in project_ids}
            for future in as_completed(futures):
                pid, stats_list, project_status = future.result()
                project_results[pid] = (stats_list, project_status)

        for project_id in project_ids:
            project = project_map[project_id]
            stats_list, project_status = project_results[project_id]

            # Build a map of completion dates from /status: {jobId: {stepName: finishDate}}
            # These finishDate values are preserved even when work is reopened
            # Use lowercase step names as keys to handle case mismatches between endpoints
            job_step_dates = {}
            if project_status:
                for job in project_status.get('jobs', []):
                    job_id = job.get('jobId')
                    if job_id:
                        job_step_dates[job_id] = {}
                        for step in job.get('steps', []):
                            step_name = step.get('workflowStepName', '')
                            finish_date = step.get('finishDate')
                            if finish_date and step_name:
                                # Store with lowercase key for case-insensitive lookup
                                job_step_dates[job_id][step_name.lower()] = finish_date

            # Process each target language
            if isinstance(stats_list, list) and stats_list:
                project_total_words = 0
                project_has_work_in_period = False
                target_languages = []
                source_lang = 'en_US'  # Default

                for lang_stats in stats_list:
                    target_lang_code = lang_stats.get('targetLanguage', 'unknown')
                    target_lang_name = self._locale_to_language_name(target_lang_code)

                    users_statistics = lang_stats.get('usersStatistics', [])

                    # Process each user's work (already filtered by get_project_statistics)
                    for user_stat in users_statistics:
                        steps_statistics = user_stat.get('stepsStatistics', [])

                        # Process each workflow step
                        for step_stat in steps_statistics:
                            step_name = step_stat.get('workflowStepName', '')
                            # Remove numbers from step name
                            clean_step_name = ''.join([c for c in step_name if not c.isdigit()])

                            jobs_statistics = step_stat.get('jobsStatistics', [])

                            # Process each job - FILTER BY STEP COMPLETION DATE
                            for job_stat in jobs_statistics:
                                include_this_job = False

                                # Get job ID to look up step finish date from /status
                                job_id = job_stat.get('jobId')

                                # Get the finish date for this specific step from /status endpoint
                                # This date is preserved even when work is reopened
                                # Use lowercase for case-insensitive lookup
                                finish_date_ts = None
                                if job_id and job_id in job_step_dates:
                                    finish_date_ts = job_step_dates[job_id].get(step_name.lower())

                                # Check if this step was completed in the target date range
                                if finish_date_ts:
                                    completion_date = datetime.fromtimestamp(finish_date_ts / 1000)
                                    if start_date <= completion_date < end_date:
                                        include_this_job = True
                                        logger.debug(f"Including job {job_id} step {step_name}: completed {completion_date}")

                                if include_this_job:
                                    # This job should be included!
                                    source_stats = job_stat.get('sourceStatistics', {})
                                    total_words = source_stats.get('totalWords', 0)

                                    if total_words > 0:
                                        project_has_work_in_period = True

                                        if target_lang_name not in target_languages:
                                            target_languages.append(target_lang_name)

                                        username = self._resolve_user_name(user_stat)

                                        # Track user statistics (user + language + workflow step)
                                        user_key = f"{username}|{target_lang_name}"
                                        if user_key not in data['user_statistics']:
                                            data['user_statistics'][user_key] = {
                                                'username': username,
                                                'language': target_lang_name,
                                                'workflow_steps': {}
                                            }

                                        # Add words to this user's workflow step
                                        if clean_step_name not in data['user_statistics'][user_key]['workflow_steps']:
                                            data['user_statistics'][user_key]['workflow_steps'][clean_step_name] = 0
                                        data['user_statistics'][user_key]['workflow_steps'][clean_step_name] += total_words

                                        # Create unique key for workflow step + language
                                        workflow_key = f"{clean_step_name} - {target_lang_name}"

                                        if workflow_key not in data['workflow_by_language']:
                                            data['workflow_by_language'][workflow_key] = {
                                                'workflow_step': clean_step_name,
                                                'language': target_lang_name,
                                                'words_done': 0,
                                                'words_to_be_done': 0,
                                                'projects': 0
                                            }

                                        data['workflow_by_language'][workflow_key]['words_done'] += total_words
                                        project_total_words += total_words

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
                        for workflow_key in data['workflow_by_language'].keys():
                            # Check if this workflow_key matches this language
                            if workflow_key.endswith(f" - {target_lang_name}"):
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

    def _get_cache_path(self, month: str) -> Path:
        """Get the JSON cache file path for a month."""
        return Path(self.config['onedrive_path']) / f".cache_monthly_{month}.json"

    def _save_month_cache(self, month: str, month_data: Dict):
        """Save processed monthly data to JSON cache."""
        cache_path = self._get_cache_path(month)
        try:
            # Extract only the serializable parts we need
            cache = {
                'languages': {},
                'users': {}
            }
            for wk, metrics in month_data.get('workflow_by_language', {}).items():
                lang = metrics['language']
                if lang not in cache['languages']:
                    cache['languages'][lang] = 0
                cache['languages'][lang] += metrics['words_done']

            for uk, ud in month_data.get('user_statistics', {}).items():
                cache['users'][uk] = {
                    'username': ud['username'],
                    'language': ud['language'],
                    'total': sum(ud['workflow_steps'].values())
                }

            with open(cache_path, 'w') as f:
                json.dump(cache, f)
            logger.info(f"Cached month data to {cache_path.name}")
        except Exception as e:
            logger.warning(f"Failed to save cache for {month}: {e}")

    def _load_month_cache(self, month: str) -> Dict:
        """Load cached monthly data. Returns None if not available."""
        cache_path = self._get_cache_path(month)
        if not cache_path.exists():
            return None
        try:
            with open(cache_path) as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"Failed to load cache for {month}: {e}")
            return None

    def aggregate_ytd_data(self, start_month: str, end_month: str, current_month_data: Dict = None) -> Dict:
        """
        Aggregate YTD data. Uses JSON cache for past months (fast),
        reuses current_month_data if provided (avoids re-querying).
        """
        logger.info(f"Aggregating YTD data from {start_month} to {end_month}")

        start_year, start_month_num = map(int, start_month.split('-'))
        end_year, end_month_num = map(int, end_month.split('-'))

        months = []
        year, month = start_year, start_month_num
        while (year < end_year) or (year == end_year and month <= end_month_num):
            months.append(f"{year}-{month:02d}")
            month += 1
            if month > 12:
                month = 1
                year += 1

        logger.info(f"Processing months: {months}")

        languages = {}
        users = {}

        for month in months:
            # If this is the current month and we already have data, reuse it
            if current_month_data and month == self.report_month:
                logger.info(f"Reusing already-queried data for {month}")
                for wk, metrics in current_month_data['workflow_by_language'].items():
                    lang = metrics['language']
                    if lang not in languages:
                        languages[lang] = {}
                    if month not in languages[lang]:
                        languages[lang][month] = 0
                    languages[lang][month] += metrics['words_done']

                for uk, ud in current_month_data['user_statistics'].items():
                    if uk not in users:
                        users[uk] = {'username': ud['username'], 'language': ud['language'], 'months': {}}
                    users[uk]['months'][month] = sum(ud['workflow_steps'].values())
                continue

            # Try JSON cache first (has proper names from previous API queries)
            cached = self._load_month_cache(month)
            if cached:
                logger.info(f"Using cached data for {month}")
                for lang, total in cached['languages'].items():
                    if lang not in languages:
                        languages[lang] = {}
                    languages[lang][month] = total

                for uk, ud in cached['users'].items():
                    if uk not in users:
                        users[uk] = {'username': ud['username'], 'language': ud['language'], 'months': {}}
                    users[uk]['months'][month] = ud['total']
                continue

            # No cache — query API
            logger.info(f"No cache for {month}, querying API...")
            try:
                month_data = self.aggregate_monthly_data(month, month)

                for wk, metrics in month_data['workflow_by_language'].items():
                    lang = metrics['language']
                    if lang not in languages:
                        languages[lang] = {}
                    if month not in languages[lang]:
                        languages[lang][month] = 0
                    languages[lang][month] += metrics['words_done']

                for uk, ud in month_data['user_statistics'].items():
                    if uk not in users:
                        users[uk] = {'username': ud['username'], 'language': ud['language'], 'months': {}}
                    users[uk]['months'][month] = sum(ud['workflow_steps'].values())

                # Save to cache for future runs
                self._save_month_cache(month, month_data)

            except Exception as e:
                logger.error(f"API query failed for {month}: {e}", exc_info=True)

        return {
            'months': months,
            'languages': languages,
            'users': users
        }

    @staticmethod
    def _generate_bar_chart_base64(labels, datasets, title, stacked=False, width=10, height=5) -> str:
        """Generate a bar chart as base64-encoded PNG."""
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import numpy as np
        import base64
        from io import BytesIO

        fig, ax = plt.subplots(figsize=(width, height))
        x = np.arange(len(labels))
        bar_width = 0.6

        if stacked and len(datasets) > 1:
            bottom = np.zeros(len(labels))
            for ds in datasets:
                ax.bar(x, ds['data'], bar_width, label=ds['label'],
                       color=ds.get('backgroundColor', '#36A2EB'), bottom=bottom)
                bottom += np.array(ds['data'])
            ax.legend(fontsize=8)
        else:
            data = datasets[0]['data'] if datasets else []
            color = datasets[0].get('backgroundColor', '#36A2EB') if datasets else '#36A2EB'
            ax.bar(x, data, bar_width, color=color)

        ax.set_title(title, fontsize=12, fontweight='bold')
        ax.set_xticks(x)
        # Truncate long labels
        short_labels = [l[:20] + '...' if len(l) > 20 else l for l in labels]
        ax.set_xticklabels(short_labels, rotation=45, ha='right', fontsize=7)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{int(v):,}'))
        ax.grid(axis='y', alpha=0.3)
        plt.tight_layout()

        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=120, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return base64.b64encode(buf.read()).decode('utf-8')

    @staticmethod
    def _generate_line_chart_base64(months, datasets, title, max_series=10, width=10, height=5) -> str:
        """Generate a line chart as base64-encoded PNG."""
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import base64
        from io import BytesIO

        fig, ax = plt.subplots(figsize=(width, height))
        colors = ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF',
                  '#FF9F40', '#E7E9ED', '#66BB6A', '#AB47BC', '#FF7043']

        for idx, ds in enumerate(datasets[:max_series]):
            ax.plot(months, ds['data'], marker='o', markersize=4,
                    label=ds['label'][:25], color=colors[idx % len(colors)], linewidth=2)

        ax.set_title(title, fontsize=12, fontweight='bold')
        ax.set_xlabel('Month', fontsize=9)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{int(v):,}'))
        ax.legend(fontsize=7, loc='upper left', bbox_to_anchor=(1.02, 1))
        ax.grid(alpha=0.3)
        plt.tight_layout()

        buf = BytesIO()
        fig.savefig(buf, format='png', dpi=120, bbox_inches='tight')
        plt.close(fig)
        buf.seek(0)
        return base64.b64encode(buf.read()).decode('utf-8')

    def create_combined_html_report(self, monthly_data: Dict, ytd_monthly_breakdown: Dict, output_path: str) -> str:
        """Create a combined interactive HTML report with both monthly and YTD data."""
        import json

        # === MONTHLY DATA ===
        workflow_by_language = monthly_data.get('workflow_by_language', {})
        user_statistics = monthly_data.get('user_statistics', {})

        # Organize monthly data by language and workflow step
        monthly_languages = {}
        for workflow_key, metrics in workflow_by_language.items():
            language = metrics['language']
            workflow_step = metrics['workflow_step']
            words = metrics['words_done']

            if language not in monthly_languages:
                monthly_languages[language] = {}
            monthly_languages[language][workflow_step] = words

        # Calculate totals per language for monthly
        monthly_language_totals = {lang: sum(steps.values()) for lang, steps in monthly_languages.items()}

        # Get all workflow steps
        all_steps = set()
        for lang_steps in monthly_languages.values():
            all_steps.update(lang_steps.keys())
        workflow_steps = sorted(all_steps, key=lambda x: ['translate', 'correct', 'final review'].index(x) if x in ['translate', 'correct', 'final review'] else 999)

        # Prepare monthly bar chart data (stacked by workflow step)
        monthly_lang_labels = sorted(monthly_language_totals.keys(), key=lambda x: monthly_language_totals[x], reverse=True)
        monthly_workflow_datasets = []
        colors = ['#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF']
        for idx, step in enumerate(workflow_steps):
            data_points = [monthly_languages.get(lang, {}).get(step, 0) for lang in monthly_lang_labels]
            monthly_workflow_datasets.append({
                'label': step,
                'data': data_points,
                'backgroundColor': colors[idx % len(colors)]
            })

        # Monthly user data for chart
        monthly_user_labels = []
        monthly_user_data_points = []
        for user_key, user_data in sorted(user_statistics.items(), key=lambda x: sum(x[1]['workflow_steps'].values()), reverse=True):
            monthly_user_labels.append(f"{user_data['username']} ({user_data['language']})")
            monthly_user_data_points.append(sum(user_data['workflow_steps'].values()))

        # Create monthly language table HTML
        monthly_lang_table_html = ""
        for language in sorted(monthly_language_totals.keys(), key=lambda x: monthly_language_totals[x], reverse=True):
            row = f"<tr><td>{language}</td>"
            for step in workflow_steps:
                row += f"<td class='number'>{monthly_languages[language].get(step, 0):,}</td>"
            row += f"<td class='number total'><strong>{monthly_language_totals[language]:,}</strong></td></tr>"
            monthly_lang_table_html += row

        # Create monthly user table HTML and build user-language mapping
        monthly_user_table_html = ""
        monthly_user_lang_map = {}  # language -> list of users
        for user_key, user_data in sorted(user_statistics.items(), key=lambda x: sum(x[1]['workflow_steps'].values()), reverse=True):
            username = user_data['username']
            language = user_data['language']

            # Build mapping
            if language not in monthly_user_lang_map:
                monthly_user_lang_map[language] = []
            if username not in monthly_user_lang_map[language]:
                monthly_user_lang_map[language].append(username)

            row = f"<tr><td>{username}</td><td>{language}</td>"
            for step in workflow_steps:
                row += f"<td class='number'>{user_data['workflow_steps'].get(step, 0):,}</td>"
            total = sum(user_data['workflow_steps'].values())
            row += f"<td class='number total'><strong>{total:,}</strong></td></tr>"
            monthly_user_table_html += row

        # === YTD DATA ===
        months = ytd_monthly_breakdown['months']
        ytd_languages = ytd_monthly_breakdown['languages']
        ytd_users = ytd_monthly_breakdown.get('users', {})

        # Prepare YTD line chart data
        ytd_lang_datasets = []
        lang_colors = [
            '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0', '#9966FF',
            '#FF9F40', '#FF6384', '#C9CBCF', '#4BC0C0', '#FF9F40'
        ]

        sorted_ytd_languages = sorted(ytd_languages.items(), key=lambda x: sum(x[1].values()), reverse=True)

        for idx, (language, month_data) in enumerate(sorted_ytd_languages):
            data_points = [month_data.get(month, 0) for month in months]
            ytd_lang_datasets.append({
                'label': language,
                'data': data_points,
                'borderColor': lang_colors[idx % len(lang_colors)],
                'backgroundColor': lang_colors[idx % len(lang_colors)] + '20',
                'tension': 0.4
            })

        # YTD user data
        ytd_user_datasets = []
        sorted_ytd_users = sorted(ytd_users.items(), key=lambda x: sum(x[1]['months'].values()), reverse=True)

        for idx, (user_key, user_data) in enumerate(sorted_ytd_users):
            username = user_data['username']
            language = user_data['language']
            data_points = [user_data['months'].get(month, 0) for month in months]
            ytd_user_datasets.append({
                'label': f"{username} ({language})",
                'data': data_points,
                'borderColor': lang_colors[idx % len(lang_colors)],
                'backgroundColor': lang_colors[idx % len(lang_colors)] + '20',
                'tension': 0.4
            })

        # Create YTD language table HTML
        ytd_lang_table_html = ""
        for language, month_data in sorted(ytd_languages.items(), key=lambda x: sum(x[1].values()), reverse=True):
            row = f"<tr><td>{language}</td>"
            for month in months:
                row += f"<td class='number'>{month_data.get(month, 0):,}</td>"
            row += f"<td class='number total'><strong>{sum(month_data.values()):,}</strong></td></tr>"
            ytd_lang_table_html += row

        # Create YTD user table HTML and build user-language mapping
        ytd_user_table_html = ""
        ytd_user_lang_map = {}  # language -> list of users
        for user_key, user_data in sorted(ytd_users.items(), key=lambda x: sum(x[1]['months'].values()), reverse=True):
            username = user_data['username']
            language = user_data['language']

            # Build mapping
            if language not in ytd_user_lang_map:
                ytd_user_lang_map[language] = []
            if username not in ytd_user_lang_map[language]:
                ytd_user_lang_map[language].append(username)

            row = f"<tr><td>{username}</td><td>{language}</td>"
            for month in months:
                row += f"<td class='number'>{user_data['months'].get(month, 0):,}</td>"
            row += f"<td class='number total'><strong>{sum(user_data['months'].values()):,}</strong></td></tr>"
            ytd_user_table_html += row

        # Generate static chart images
        logger.info("Generating chart images...")
        monthly_lang_chart_b64 = self._generate_bar_chart_base64(
            monthly_lang_labels, monthly_workflow_datasets,
            'Words Processed by Language and Workflow Step', stacked=True)

        monthly_user_ds = [{'label': 'Total Words', 'data': monthly_user_data_points[:20],
                            'backgroundColor': '#36A2EB'}]
        monthly_user_chart_b64 = self._generate_bar_chart_base64(
            [l[:30] for l in monthly_user_labels[:20]], monthly_user_ds,
            'Words Processed by User')

        ytd_lang_ds = [{'label': l, 'data': [ytd_languages[l].get(m, 0) for m in months]}
                       for l, _ in sorted_ytd_languages[:10]]
        ytd_lang_chart_b64 = self._generate_line_chart_base64(
            months, ytd_lang_ds, 'Language Trends (Top 10)')

        ytd_user_ds = [{'label': f"{ud['username']} ({ud['language']})",
                        'data': [ud['months'].get(m, 0) for m in months]}
                       for _, ud in sorted_ytd_users[:10]]
        ytd_user_chart_b64 = self._generate_line_chart_base64(
            months, ytd_user_ds, 'User Productivity (Top 10)')

        html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>XTM Report - {self.report_month}</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; color: #333; }}
        .container {{ max-width: 1600px; margin: 0 auto; }}
        .header {{ background: white; padding: 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }}
        h1 {{ color: #366092; font-size: 2em; margin-bottom: 10px; }}
        .subtitle {{ color: #666; font-size: 1.1em; }}
        .section-divider {{ background: white; padding: 20px 30px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin: 30px 0 20px 0; }}
        .section-divider h2 {{ color: #366092; margin: 0; font-size: 1.8em; }}
        .filter-panel {{ background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin-bottom: 20px; }}
        .filter-group {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
        .filter-item {{ display: flex; flex-direction: column; }}
        .filter-item label {{ font-weight: 600; color: #366092; margin-bottom: 8px; font-size: 1.1em; }}
        .filter-item input[type="text"] {{ padding: 8px; border: 2px solid #ddd; border-radius: 5px; font-size: 16px; margin-bottom: 10px; }}
        .filter-item input[type="text"]:focus {{ outline: none; border-color: #366092; }}
        .checkbox-list {{ max-height: 200px; overflow-y: auto; border: 1px solid #ddd; border-radius: 5px; padding: 10px; background: #fafafa; }}
        .checkbox-item {{ display: flex; align-items: center; padding: 5px 0; }}
        .checkbox-item input {{ margin-right: 8px; cursor: pointer; }}
        .checkbox-item label {{ cursor: pointer; margin: 0; font-weight: normal; color: #333; }}
        .data-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }}
        .data-card {{ background: white; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); overflow: hidden; }}
        h3 {{ color: #366092; margin-bottom: 15px; padding-bottom: 8px; border-bottom: 3px solid #366092; font-size: 1.2em; }}
        .chart-img {{ width: 100%; height: auto; display: block; }}
        .chart-container {{ position: relative; height: 400px; display: none; }}
        @media (max-width: 768px) {{
            .chart-container {{ height: 250px; }}
        }}
        .table-container {{ overflow-x: auto; max-height: 400px; overflow-y: auto; -webkit-overflow-scrolling: touch; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 0.85em; }}
        th {{ background: #366092; color: white; padding: 8px; text-align: left; font-weight: 600; cursor: pointer; user-select: none; position: sticky; top: 0; z-index: 10; white-space: nowrap; }}
        th:hover {{ background: #2a4d73; }}
        th.sortable:after {{ content: ' ⇅'; opacity: 0.5; font-size: 0.8em; }}
        th.sort-asc:after {{ content: ' ▲'; opacity: 1; }}
        th.sort-desc:after {{ content: ' ▼'; opacity: 1; }}
        td {{ padding: 6px 8px; border-bottom: 1px solid #e0e0e0; white-space: nowrap; }}
        tr:hover {{ background: #f5f5f5; }}
        .number {{ text-align: right; }}
        .total {{ background: #f0f7ff; }}
        @media (max-width: 1400px) {{
            .data-row {{ grid-template-columns: 1fr; }}
            .filter-group {{ grid-template-columns: 1fr; }}
        }}
        @media (max-width: 768px) {{
            body {{ padding: 8px; }}
            .header {{ padding: 15px; }}
            h1 {{ font-size: 1.4em; }}
            .subtitle {{ font-size: 0.9em; }}
            .section-divider {{ padding: 12px 15px; }}
            .section-divider h2 {{ font-size: 1.3em; }}
            .filter-panel {{ padding: 12px; }}
            .data-card {{ padding: 10px; }}
            h3 {{ font-size: 1em; }}
            .checkbox-list {{ max-height: 150px; }}
            table {{ font-size: 0.75em; }}
            th, td {{ padding: 4px 6px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 XTM Translation Report</h1>
            <div class="subtitle">
                <strong>Report Period:</strong> {self.report_month_name} {self.report_month.split('-')[0]}<br>
                <strong>Year-to-Date:</strong> {self.ytd_start_month} to {self.ytd_end_month}<br>
                <strong>Generated:</strong> {self.report_date.strftime('%Y-%m-%d %H:%M')}
            </div>
        </div>

        <!-- MONTHLY SECTION -->
        <div class="section-divider">
            <h2>📅 Monthly Report - {self.report_month_name}</h2>
        </div>

        <div class="filter-panel">
            <h3>🔍 Filters</h3>
            <div class="filter-group">
                <div class="filter-item">
                    <label>Languages:</label>
                    <input type="text" id="monthlyLangSearch" placeholder="Type to search..." onkeyup="filterCheckboxes('monthlyLangCheckboxes', this.value)">
                    <div class="checkbox-list" id="monthlyLangCheckboxes">
                        {''.join(['<div class="checkbox-item"><input type="checkbox" id="monthlyLang_' + lang.replace(" ", "_") + '" value="' + lang + '" onchange="applyFilters(' + "'monthly'" + ')"><label for="monthlyLang_' + lang.replace(" ", "_") + '">' + lang + '</label></div>' for lang in sorted(monthly_language_totals.keys())])}
                    </div>
                </div>
                <div class="filter-item">
                    <label>Users:</label>
                    <input type="text" id="monthlyUserSearch" placeholder="Type to search..." onkeyup="filterCheckboxes('monthlyUserCheckboxes', this.value)">
                    <div class="checkbox-list" id="monthlyUserCheckboxes">
                        {''.join(['<div class="checkbox-item" data-lang="' + user_data["language"] + '"><input type="checkbox" id="monthlyUser_' + user_data["username"].replace(" ", "_") + '" value="' + user_data["username"] + '" onchange="applyFilters(' + "'monthly'" + ')"><label for="monthlyUser_' + user_data["username"].replace(" ", "_") + '">' + user_data["username"] + '</label></div>' for user_key, user_data in sorted(user_statistics.items(), key=lambda x: x[1]["username"])])}
                    </div>
                </div>
            </div>
        </div>

        <div class="data-row">
            <div class="data-card">
                <h3>🌍 Language Translation Volume</h3>
                <img class="chart-img" id="monthlyLangImg" src="data:image/png;base64,{monthly_lang_chart_b64}" alt="Language Chart">
                <div class="chart-container" id="monthlyLangChartWrap"><canvas id="monthlyLanguageChart"></canvas></div>
            </div>
            <div class="data-card">
                <h3>📊 Language Data</h3>
                <div class="table-container">
                    <table id="monthlyLangTable">
                        <thead>
                            <tr>
                                <th class="sortable" onclick="sortTable('monthlyLangTable', 0)">Language</th>
                                {''.join([f'<th class="number sortable" onclick="sortTable(' + "'monthlyLangTable', " + str(idx+1) + ')">' + step + '</th>' for idx, step in enumerate(workflow_steps)])}
                                <th class="number sortable" onclick="sortTable('monthlyLangTable', {len(workflow_steps)+1})">Total</th>
                            </tr>
                        </thead>
                        <tbody>
                            {monthly_lang_table_html}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <div class="data-row">
            <div class="data-card">
                <h3>👥 User Productivity</h3>
                <img class="chart-img" id="monthlyUserImg" src="data:image/png;base64,{monthly_user_chart_b64}" alt="User Chart">
                <div class="chart-container" id="monthlyUserChartWrap"><canvas id="monthlyUserChart"></canvas></div>
            </div>
            <div class="data-card">
                <h3>📊 User Data</h3>
                <div class="table-container">
                    <table id="monthlyUserTable">
                        <thead>
                            <tr>
                                <th class="sortable" onclick="sortTable('monthlyUserTable', 0)">Name</th>
                                <th class="sortable" onclick="sortTable('monthlyUserTable', 1)">Language</th>
                                {''.join([f'<th class="number sortable" onclick="sortTable(' + "'monthlyUserTable', " + str(idx+2) + ')">' + step + '</th>' for idx, step in enumerate(workflow_steps)])}
                                <th class="number sortable" onclick="sortTable('monthlyUserTable', {len(workflow_steps)+2})">Total</th>
                            </tr>
                        </thead>
                        <tbody>
                            {monthly_user_table_html}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <!-- YTD SECTION -->
        <div class="section-divider">
            <h2>📈 Year-to-Date Report - {self.ytd_start_month} to {self.ytd_end_month}</h2>
        </div>

        <div class="filter-panel">
            <h3>🔍 Filters</h3>
            <div class="filter-group">
                <div class="filter-item">
                    <label>Languages:</label>
                    <input type="text" id="ytdLangSearch" placeholder="Type to search..." onkeyup="filterCheckboxes('ytdLangCheckboxes', this.value)">
                    <div class="checkbox-list" id="ytdLangCheckboxes">
                        {''.join(['<div class="checkbox-item"><input type="checkbox" id="ytdLang_' + lang.replace(" ", "_") + '" value="' + lang + '" onchange="applyFilters(' + "'ytd'" + ')"><label for="ytdLang_' + lang.replace(" ", "_") + '">' + lang + '</label></div>' for lang in sorted(ytd_languages.keys())])}
                    </div>
                </div>
                <div class="filter-item">
                    <label>Users:</label>
                    <input type="text" id="ytdUserSearch" placeholder="Type to search..." onkeyup="filterCheckboxes('ytdUserCheckboxes', this.value)">
                    <div class="checkbox-list" id="ytdUserCheckboxes">
                        {''.join(['<div class="checkbox-item" data-lang="' + user_data["language"] + '"><input type="checkbox" id="ytdUser_' + user_data["username"].replace(" ", "_") + '" value="' + user_data["username"] + '" onchange="applyFilters(' + "'ytd'" + ')"><label for="ytdUser_' + user_data["username"].replace(" ", "_") + '">' + user_data["username"] + '</label></div>' for user_key, user_data in sorted(ytd_users.items(), key=lambda x: x[1]["username"])])}
                    </div>
                </div>
            </div>
        </div>

        <div class="data-row">
            <div class="data-card">
                <h3>🌍 Language Translation Trends</h3>
                <img class="chart-img" id="ytdLangImg" src="data:image/png;base64,{ytd_lang_chart_b64}" alt="YTD Language Chart">
                <div class="chart-container" id="ytdLangChartWrap"><canvas id="ytdLanguageChart"></canvas></div>
            </div>
            <div class="data-card">
                <h3>📊 Language Data</h3>
                <div class="table-container">
                    <table id="ytdLangTable">
                        <thead>
                            <tr>
                                <th class="sortable" onclick="sortTable('ytdLangTable', 0)">Language</th>
                                {''.join([f'<th class="number sortable" onclick="sortTable(' + "'ytdLangTable', " + str(idx+1) + ')">' + month + '</th>' for idx, month in enumerate(months)])}
                                <th class="number sortable" onclick="sortTable('ytdLangTable', {len(months)+1})">Total</th>
                            </tr>
                        </thead>
                        <tbody>
                            {ytd_lang_table_html}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>

        <div class="data-row">
            <div class="data-card">
                <h3>👥 User Productivity Trends</h3>
                <img class="chart-img" id="ytdUserImg" src="data:image/png;base64,{ytd_user_chart_b64}" alt="YTD User Chart">
                <div class="chart-container" id="ytdUserChartWrap"><canvas id="ytdUserChart"></canvas></div>
            </div>
            <div class="data-card">
                <h3>📊 User Data</h3>
                <div class="table-container">
                    <table id="ytdUserTable">
                        <thead>
                            <tr>
                                <th class="sortable" onclick="sortTable('ytdUserTable', 0)">Name</th>
                                <th class="sortable" onclick="sortTable('ytdUserTable', 1)">Language</th>
                                {''.join([f'<th class="number sortable" onclick="sortTable(' + "'ytdUserTable', " + str(idx+2) + ')">' + month + '</th>' for idx, month in enumerate(months)])}
                                <th class="number sortable" onclick="sortTable('ytdUserTable', {len(months)+2})">Total</th>
                            </tr>
                        </thead>
                        <tbody>
                            {ytd_user_table_html}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>

    <script>
        // User-language mappings
        const monthlyUserLangMap = {json.dumps(monthly_user_lang_map)};
        const ytdUserLangMap = {json.dumps(ytd_user_lang_map)};

        // Initialize Chart.js if available (swap static images for dynamic canvases)
        let monthlyLanguageChart, monthlyUserChart, ytdLanguageChart, ytdUserChart;
        if (typeof Chart !== 'undefined') {{
            // Hide static images, show canvases
            ['monthlyLangImg','monthlyUserImg','ytdLangImg','ytdUserImg'].forEach(id => {{
                const el = document.getElementById(id);
                if (el) el.style.display = 'none';
            }});
            ['monthlyLangChartWrap','monthlyUserChartWrap','ytdLangChartWrap','ytdUserChartWrap'].forEach(id => {{
                const el = document.getElementById(id);
                if (el) el.style.display = 'block';
            }});

            const fmtTick = v => v.toLocaleString();

            monthlyLanguageChart = new Chart(document.getElementById('monthlyLanguageChart'), {{
                type: 'bar',
                data: {{ labels: {json.dumps(monthly_lang_labels)}, datasets: {json.dumps(monthly_workflow_datasets)} }},
                options: {{ responsive: true, maintainAspectRatio: false,
                    plugins: {{ title: {{ display: true, text: 'Words by Language & Workflow Step' }}, legend: {{ position: 'top' }} }},
                    scales: {{ x: {{ stacked: true }}, y: {{ stacked: true, beginAtZero: true, ticks: {{ callback: fmtTick }} }} }}
                }}
            }});

            monthlyUserChart = new Chart(document.getElementById('monthlyUserChart'), {{
                type: 'bar',
                data: {{ labels: {json.dumps(monthly_user_labels)}, datasets: [{{ label: 'Total Words', data: {json.dumps(monthly_user_data_points)}, backgroundColor: '#36A2EB' }}] }},
                options: {{ responsive: true, maintainAspectRatio: false,
                    plugins: {{ title: {{ display: true, text: 'Words by User' }}, legend: {{ display: false }} }},
                    scales: {{ y: {{ beginAtZero: true, ticks: {{ callback: fmtTick }} }} }}
                }}
            }});

            ytdLanguageChart = new Chart(document.getElementById('ytdLanguageChart'), {{
                type: 'line',
                data: {{ labels: {json.dumps(months)}, datasets: {json.dumps(ytd_lang_datasets)} }},
                options: {{ responsive: true, maintainAspectRatio: false,
                    plugins: {{ title: {{ display: true, text: 'Language Trends' }}, legend: {{ position: 'right', labels: {{ boxWidth: 12, font: {{ size: 9 }} }} }} }},
                    scales: {{ y: {{ beginAtZero: true, ticks: {{ callback: fmtTick }} }} }}
                }}
            }});

            ytdUserChart = new Chart(document.getElementById('ytdUserChart'), {{
                type: 'line',
                data: {{ labels: {json.dumps(months)}, datasets: {json.dumps(ytd_user_datasets)} }},
                options: {{ responsive: true, maintainAspectRatio: false,
                    plugins: {{ title: {{ display: true, text: 'User Productivity' }}, legend: {{ position: 'right', labels: {{ boxWidth: 12, font: {{ size: 9 }} }} }} }},
                    scales: {{ y: {{ beginAtZero: true, ticks: {{ callback: fmtTick }} }} }}
                }}
            }});
        }}

        function updateChartFromTable(tableId) {{
            if (typeof Chart === 'undefined') return;
            const table = document.getElementById(tableId);
            const rows = table.querySelectorAll('tbody tr');
            const visibleLabels = [];
            rows.forEach(row => {{ if (row.style.display !== 'none') visibleLabels.push(row.cells[0].textContent.trim()); }});

            if (tableId === 'monthlyLangTable') updateBarChart(monthlyLanguageChart, {json.dumps(monthly_lang_labels)}, {json.dumps(monthly_workflow_datasets)}, visibleLabels);
            else if (tableId === 'monthlyUserTable') updateBarChart(monthlyUserChart, {json.dumps(monthly_user_labels)}, [{{ label: 'Total Words', data: {json.dumps(monthly_user_data_points)}, backgroundColor: '#36A2EB' }}], visibleLabels);
            else if (tableId === 'ytdLangTable') updateLineChart(ytdLanguageChart, {json.dumps(ytd_lang_datasets)}, visibleLabels);
            else if (tableId === 'ytdUserTable') updateLineChart(ytdUserChart, {json.dumps(ytd_user_datasets)}, visibleLabels);
        }}

        function updateBarChart(chart, allLabels, allDatasets, visibleLabels) {{
            if (!chart) return;
            const idx = []; allLabels.forEach((l, i) => {{ if (visibleLabels.some(v => l.includes(v) || v.includes(l))) idx.push(i); }});
            chart.data.labels = idx.map(i => allLabels[i]);
            chart.data.datasets = allDatasets.map(ds => ({{ ...ds, data: idx.map(i => ds.data[i]) }}));
            chart.update();
        }}

        function updateLineChart(chart, allDatasets, visibleLabels) {{
            if (!chart) return;
            chart.data.datasets = allDatasets.filter(ds => {{ const n = ds.label.split(' (')[0]; return visibleLabels.some(v => v === ds.label || v === n || ds.label.includes(v)); }});
            chart.update();
        }}

        function filterCheckboxes(containerId, searchValue) {{
            const container = document.getElementById(containerId);
            const items = container.querySelectorAll('.checkbox-item');
            const search = searchValue.toUpperCase();

            items.forEach(item => {{
                const label = item.querySelector('label').textContent.toUpperCase();
                item.style.display = label.indexOf(search) > -1 ? '' : 'none';
            }});
        }}

        function applyFilters(section) {{
            // Get checked language checkboxes
            const langCheckboxes = document.querySelectorAll(`#${{section}}LangCheckboxes input[type="checkbox"]:checked`);
            const selectedLangs = Array.from(langCheckboxes).map(cb => cb.value.toUpperCase());

            // Get checked user checkboxes
            const userCheckboxes = document.querySelectorAll(`#${{section}}UserCheckboxes input[type="checkbox"]:checked`);
            const selectedUsers = Array.from(userCheckboxes).map(cb => cb.value.toUpperCase());

            // Update user checkboxes based on selected languages
            updateUserCheckboxes(section, selectedLangs);

            // Filter language table
            const langTable = document.getElementById(`${{section}}LangTable`);
            const langRows = langTable.querySelectorAll('tbody tr');
            langRows.forEach(row => {{
                const langName = row.cells[0].textContent.trim().toUpperCase();
                const show = selectedLangs.length === 0 || selectedLangs.includes(langName);
                row.style.display = show ? '' : 'none';
            }});

            // Filter user table
            const userTable = document.getElementById(`${{section}}UserTable`);
            const userRows = userTable.querySelectorAll('tbody tr');
            userRows.forEach(row => {{
                const userName = row.cells[0].textContent.trim().toUpperCase();
                const userLang = row.cells[1].textContent.trim().toUpperCase();
                let show = true;

                // Apply user checkbox filter
                if (selectedUsers.length > 0 && !selectedUsers.includes(userName)) {{
                    show = false;
                }}

                // Apply language filter to user table
                if (selectedLangs.length > 0 && !selectedLangs.includes(userLang)) {{
                    show = false;
                }}

                row.style.display = show ? '' : 'none';
            }});

            // Update charts
            updateChartFromTable(`${{section}}LangTable`);
            updateChartFromTable(`${{section}}UserTable`);
        }}

        function updateUserCheckboxes(section, selectedLangs) {{
            const userContainer = document.getElementById(`${{section}}UserCheckboxes`);
            const userItems = userContainer.querySelectorAll('.checkbox-item');
            const userLangMap = section === 'monthly' ? monthlyUserLangMap : ytdUserLangMap;
            const searchInput = document.getElementById(`${{section}}UserSearch`);
            const searchVal = searchInput ? searchInput.value.toUpperCase() : '';

            // Build set of users available for selected languages
            const availableUsers = new Set();
            if (selectedLangs.length > 0) {{
                selectedLangs.forEach(langUpper => {{
                    Object.keys(userLangMap).forEach(mapLang => {{
                        if (mapLang.toUpperCase() === langUpper) {{
                            userLangMap[mapLang].forEach(user => availableUsers.add(user.toUpperCase()));
                        }}
                    }});
                }});
            }}

            userItems.forEach(item => {{
                const checkbox = item.querySelector('input[type="checkbox"]');
                const userName = checkbox.value.toUpperCase();
                const labelText = item.querySelector('label').textContent.toUpperCase();

                // Must pass both: language filter AND search filter
                const passesLang = selectedLangs.length === 0 || availableUsers.has(userName);
                const passesSearch = searchVal === '' || labelText.indexOf(searchVal) > -1;

                if (passesLang && passesSearch) {{
                    item.style.display = '';
                }} else {{
                    item.style.display = 'none';
                    if (!passesLang) {{
                        checkbox.checked = false;
                    }}
                }}
            }});
        }}

        function sortTable(tableId, columnIndex) {{
            const table = document.getElementById(tableId);
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            const th = table.querySelectorAll('th')[columnIndex];

            const currentSort = th.classList.contains('sort-asc') ? 'asc' :
                               th.classList.contains('sort-desc') ? 'desc' : 'none';
            const newSort = currentSort === 'none' ? 'desc' :
                           currentSort === 'desc' ? 'asc' : 'desc';

            table.querySelectorAll('th').forEach(header => {{
                header.classList.remove('sort-asc', 'sort-desc');
            }});

            if (newSort === 'asc') {{
                th.classList.add('sort-asc');
            }} else {{
                th.classList.add('sort-desc');
            }}

            rows.sort((a, b) => {{
                const aCell = a.cells[columnIndex];
                const bCell = b.cells[columnIndex];

                let aVal = aCell.textContent.trim();
                let bVal = bCell.textContent.trim();

                aVal = aVal.replace(/,/g, '');
                bVal = bVal.replace(/,/g, '');

                const aNum = parseFloat(aVal);
                const bNum = parseFloat(bVal);

                let comparison = 0;
                if (!isNaN(aNum) && !isNaN(bNum)) {{
                    comparison = aNum - bNum;
                }} else {{
                    comparison = aVal.localeCompare(bVal);
                }}

                return newSort === 'asc' ? comparison : -comparison;
            }});

            rows.forEach(row => tbody.appendChild(row));
            updateChartFromTable(tableId);
        }}
    </script>
</body>
</html>"""

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

        return output_path

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

The attached HTML report is interactive and includes:

Monthly Section:
• Bar charts showing translation volume by language and workflow step
• User productivity bar charts

Year-to-Date Section:
• Line charts showing monthly trends for all languages
• User productivity trends over time

All sections include:
• Sortable tables (click any column header to sort)
• Filterable data (use the search boxes to filter by language or user)

Report Generated: {self.report_date.strftime('%Y-%m-%d %H:%M')}

Please open the HTML file in your web browser to view the complete interactive report.

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
        """Main method to generate and distribute the HTML report."""
        html_path = None
        try:
            logger.info("Starting XTM monthly report generation")

            # Run health checks
            self._run_health_checks()

            # Query API for current month data
            try:
                monthly_data = self.aggregate_monthly_data(self.report_month, self.report_month)
                # Cache for future runs
                self._save_month_cache(self.report_month, monthly_data)
            except Exception as e:
                logger.error(f"Failed to aggregate monthly data: {e}", exc_info=True)
                monthly_data = {
                    'project_stats': {'total': 0, 'completed': 0, 'in_progress': 0, 'pending': 0},
                    'workflow_by_language': {},
                    'user_statistics': {},
                    'projects': []
                }

            # Aggregate YTD data (reuses monthly_data for current month, cache for past months)
            try:
                ytd_monthly_breakdown = self.aggregate_ytd_data(
                    self.ytd_start_month, self.ytd_end_month,
                    current_month_data=monthly_data
                )
            except Exception as e:
                logger.error(f"Failed to aggregate YTD data: {e}", exc_info=True)
                ytd_monthly_breakdown = {'months': [], 'languages': {}, 'users': {}}

            if not (monthly_data.get('workflow_by_language') or ytd_monthly_breakdown.get('languages')):
                logger.warning("No data available for reporting period")

            # Ensure output directory exists
            output_dir = Path(self.config['onedrive_path'])
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                logger.warning(f"Could not create output directory: {e}")

            # Create HTML report
            html_filename = f"XTM_Report_{self.report_month}_{self.report_date.strftime('%Y%m%d')}.html"
            html_path = str(output_dir / html_filename)
            self.create_combined_html_report(monthly_data, ytd_monthly_breakdown, html_path)
            logger.info(f"Combined HTML report saved to {html_path}")

            # Construct a minimal ytd_data for the email body stats
            ytd_data = {
                'project_stats': {'total': 0, 'completed': 0, 'in_progress': 0, 'pending': 0},
                'workflow_by_language': {},
                'user_statistics': {},
                'projects': []
            }
            # Populate from YTD breakdown for email summary
            for language, month_words in ytd_monthly_breakdown.get('languages', {}).items():
                total = sum(month_words.values())
                ytd_data['workflow_by_language'][f"translate - {language}"] = {
                    'workflow_step': 'translate', 'language': language,
                    'words_done': total, 'words_to_be_done': 0, 'projects': 0
                }

            # Send via Outlook
            try:
                self.send_email_via_outlook(html_path, monthly_data, ytd_data)
                email_success = True
            except Exception as e:
                logger.error(f"Failed to send email: {e}", exc_info=True)
                email_success = False

            logger.info("Report generation completed successfully")
            print(f"\n✓ HTML Report generated: {html_path}")
            print(f"✓ Period: {self.report_month} (YTD: {self.ytd_start_month} to {self.ytd_end_month})")
            if email_success:
                if self.auto_send:
                    print(f"✓ Email sent automatically to {len(self.config['email_recipients'])} recipients")
                else:
                    print(f"✓ Email draft created with {len(self.config['email_recipients'])} recipients")
            else:
                print(f"⚠ Email could not be sent - report saved locally")

        except Exception as e:
            logger.error(f"Report generation failed: {e}", exc_info=True)
            try:
                self._send_failure_notification(str(e), html_path)
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
