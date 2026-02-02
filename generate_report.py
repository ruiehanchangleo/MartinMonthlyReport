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
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Any
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
            with open(config_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Failed to load config: {e}")
            raise

    def _locale_to_language_name(self, locale_code: str) -> str:
        """Convert locale code to language name."""
        return self.LOCALE_TO_LANGUAGE.get(locale_code, locale_code)

    def _make_request(self, endpoint: str, method: str = 'GET', params: Dict = None, data: Dict = None) -> Any:
        """Make API request to XTM with error handling."""
        url = f"{self.base_url}/{endpoint}"
        try:
            logger.info(f"Making {method} request to {endpoint}")
            if method == 'GET':
                response = requests.get(url, headers=self.headers, params=params, timeout=30)
            elif method == 'POST':
                response = requests.post(url, headers=self.headers, json=data, timeout=30)
            else:
                raise ValueError(f"Unsupported method: {method}")

            response.raise_for_status()
            return response.json() if response.content else {}
        except requests.exceptions.RequestException as e:
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

            # Filter projects by month (if date available)
            project_date = project.get('createdDate') or project.get('modificationDate')
            readable_date = None
            project_month = None

            if project_date:
                try:
                    if isinstance(project_date, (int, float)):
                        dt = datetime.fromtimestamp(project_date / 1000)
                        readable_date = dt.strftime('%Y-%m-%d')
                        project_month = dt.strftime('%Y-%m')
                except:
                    readable_date = str(project_date)

            # Skip projects outside the date range (only if we have a valid date)
            # If no date is available, include the project
            if project_month:
                if project_month < start_month or project_month > end_month:
                    continue
            # If no date available, include all projects (they might still have metrics)

            # Update project stats
            data['project_stats']['total'] += 1
            status = project.get('status', 'UNKNOWN')

            if status == 'FINISHED':
                data['project_stats']['completed'] += 1
            elif status in ['IN_PROGRESS', 'STARTED']:
                data['project_stats']['in_progress'] += 1
            else:
                data['project_stats']['pending'] += 1

            # Get project statistics (filtered to exclude leo.chang@familysearch.org)
            stats_list = self.get_project_statistics(project_id)

            # Process each target language in the statistics
            if isinstance(stats_list, list) and stats_list:
                project_total_words = 0
                target_languages = []
                source_lang = 'en_US'  # Default, will try to determine from context

                for lang_stats in stats_list:
                    target_lang_code = lang_stats.get('targetLanguage', 'unknown')
                    target_lang_name = self._locale_to_language_name(target_lang_code)
                    target_languages.append(target_lang_name)

                    # Aggregate words from all users (already filtered to exclude leo.chang@familysearch.org)
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
                                # Sum words from different metrics (source statistics total words)
                                source_stats = job_stat.get('sourceStatistics', {})
                                step_total_words += source_stats.get('totalWords', 0)

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

                    # Mark this project as counted for this language
                    for step_name in set([s.get('workflowStepName', '') for u in users_statistics for s in u.get('stepsStatistics', [])]):
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
                    'created_date': readable_date
                })
            else:
                # No metrics available, still add project
                data['projects'].append({
                    'id': project_id,
                    'name': project.get('name', 'Unknown'),
                    'status': status,
                    'source_lang': 'unknown',
                    'target_langs': 'N/A',
                    'total_words': 0,
                    'created_date': readable_date
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

        # Save workbook
        wb.save(output_path)
        logger.info(f"Excel report saved to {output_path}")

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
        try:
            logger.info("Starting XTM monthly report generation")

            # Aggregate monthly data
            monthly_data = self.aggregate_monthly_data(self.report_month, self.report_month)

            # Aggregate year-to-date data
            ytd_data = self.aggregate_monthly_data(self.ytd_start_month, self.ytd_end_month)

            # Create output filename
            output_filename = f"XTM_Monthly_Report_{self.report_month}_{self.report_date.strftime('%Y%m%d')}.xlsx"
            output_path = Path(self.config['onedrive_path']) / output_filename

            # Ensure output directory exists
            output_path.parent.mkdir(parents=True, exist_ok=True)

            # Create Excel report
            report_path = self.create_excel_report(monthly_data, ytd_data, str(output_path))

            # Send via Outlook (pass the data for email body)
            self.send_email_via_outlook(report_path, monthly_data, ytd_data)

            logger.info("Report generation completed successfully")
            print(f"\n✓ Report generated: {report_path}")
            print(f"✓ Monthly period: {self.report_month_name}")
            print(f"✓ YTD period: {self.ytd_start_month} to {self.ytd_end_month}")
            if self.auto_send:
                print(f"✓ Email sent automatically to {len(self.config['email_recipients'])} recipients")
            else:
                print(f"✓ Email draft created with {len(self.config['email_recipients'])} recipients")

        except Exception as e:
            logger.error(f"Report generation failed: {e}", exc_info=True)
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
