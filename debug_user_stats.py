#!/usr/bin/env python3
"""
Debug script to show detailed user statistics for all projects.
Shows who worked on what, including the excluded user.
"""

import json
import sys
from generate_report import XTMReportGenerator

def main():
    # Create report generator
    generator = XTMReportGenerator('xtm_config.json', auto_send=False)

    print("=" * 80)
    print("XTM USER STATISTICS DEBUG REPORT")
    print("=" * 80)
    print(f"\nExcluded users: leo.chang@familysearch.org, LeoAdmin")
    print(f"Report period: {generator.report_month}")
    print("\n")

    # Get all projects
    projects = generator.get_projects()
    print(f"Total projects: {len(projects)}\n")

    excluded_user_found = False
    projects_with_excluded = []
    total_users_by_email = {}

    # Process each project
    for idx, project in enumerate(projects, 1):
        project_id = project['id']
        project_name = project['name']

        # Get statistics for this project (pass empty list to not exclude anyone)
        stats_list = generator.get_project_statistics(project_id, excluded_users=[])

        if not stats_list:
            continue

        project_has_excluded = False
        project_users = {}

        # Process each target language
        for lang_stats in stats_list:
            target_language = lang_stats.get('targetLanguage', 'Unknown')
            # Convert locale to language name
            language_name = generator._locale_to_language_name(target_language)
            users_statistics = lang_stats.get('usersStatistics', [])

            for user_stats in users_statistics:
                username = user_stats.get('username', 'Unknown')

                # Track if this is an excluded user
                is_excluded = username.lower() in ['leo.chang@familysearch.org', 'leoadmin']
                if is_excluded:
                    excluded_user_found = True
                    project_has_excluded = True

                # Count users
                if username not in total_users_by_email:
                    total_users_by_email[username] = {
                        'projects': set(),
                        'total_words': 0,
                        'is_excluded': is_excluded,
                        'languages': {}
                    }

                total_users_by_email[username]['projects'].add(project_id)

                # Get workflow steps
                steps_statistics = user_stats.get('stepsStatistics', [])

                for step_stat in steps_statistics:
                    step_name = step_stat.get('workflowStepName', '')
                    jobs_statistics = step_stat.get('jobsStatistics', [])

                    step_total_words = 0
                    for job_stat in jobs_statistics:
                        source_stats = job_stat.get('sourceStatistics', {})
                        words = source_stats.get('totalWords', 0)
                        step_total_words += words

                    if step_total_words > 0:
                        if username not in project_users:
                            project_users[username] = {
                                'is_excluded': is_excluded,
                                'languages': {}
                            }

                        if language_name not in project_users[username]['languages']:
                            project_users[username]['languages'][language_name] = {}

                        project_users[username]['languages'][language_name][step_name] = step_total_words
                        total_users_by_email[username]['total_words'] += step_total_words

                        # Track language totals for the user
                        if language_name not in total_users_by_email[username]['languages']:
                            total_users_by_email[username]['languages'][language_name] = 0
                        total_users_by_email[username]['languages'][language_name] += step_total_words

        # Print project details if it has users
        if project_users:
            if project_has_excluded:
                projects_with_excluded.append(project_name)

            print(f"\n{'─' * 80}")
            print(f"Project #{idx}: {project_name} (ID: {project_id})")
            print(f"{'─' * 80}")

            for username, user_data in sorted(project_users.items()):
                excluded_marker = " ⚠️  EXCLUDED USER" if user_data['is_excluded'] else ""
                print(f"\n  👤 User: {username}{excluded_marker}")

                for lang, steps in sorted(user_data['languages'].items()):
                    print(f"     Language: {lang}")
                    for step, words in sorted(steps.items()):
                        print(f"       • {step}: {words:,} words")

    # Summary
    print("\n\n")
    print("=" * 80)
    print("SUMMARY")
    print("=" * 80)

    print(f"\nTotal unique users: {len(total_users_by_email)}")
    print(f"\nAll users found:")
    for username, data in sorted(total_users_by_email.items()):
        excluded_marker = " ⚠️  EXCLUDED" if data['is_excluded'] else ""
        print(f"  • {username}{excluded_marker}")
        print(f"    - Projects: {len(data['projects'])}")
        print(f"    - Total words: {data['total_words']:,}")

        # Show top languages for this user
        if data['languages']:
            sorted_langs = sorted(data['languages'].items(), key=lambda x: x[1], reverse=True)
            print(f"    - Languages ({len(sorted_langs)}):")
            for lang, words in sorted_langs[:5]:  # Show top 5 languages
                print(f"      • {lang}: {words:,} words")
            if len(sorted_langs) > 5:
                print(f"      • ... and {len(sorted_langs) - 5} more")

    if excluded_user_found:
        print(f"\n⚠️  EXCLUDED USERS FOUND!")
        print(f"   Excluded users (leo.chang@familysearch.org or LeoAdmin) worked on {len(projects_with_excluded)} projects:")
        for proj in projects_with_excluded:
            print(f"   - {proj}")
    else:
        print(f"\n✓ Excluded users (leo.chang@familysearch.org, LeoAdmin) NOT found in any projects")

    print("\n" + "=" * 80)

if __name__ == '__main__':
    main()
