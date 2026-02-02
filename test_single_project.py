#!/usr/bin/env python3
"""Test parsing a single project to debug"""

import json
import requests

# Load config
with open('xtm_config.json', 'r') as f:
    config = json.load(f)

base_url = config['base_url']
headers = {
    'Authorization': f"{config['auth_type']} {config['auth_token']}",
    'Content-Type': 'application/json'
}

# Get first project
response = requests.get(f"{base_url}/projects", headers=headers)
projects = response.json()
project = projects[0]
project_id = project['id']

print(f"Project: {project.get('name')}")
print(f"Status: {project.get('status')}")

# Get metrics
metrics_response = requests.get(f"{base_url}/projects/{project_id}/metrics", headers=headers)
metrics_list = metrics_response.json()

print(f"\nMetrics is a list: {isinstance(metrics_list, list)}")
print(f"Metrics length: {len(metrics_list) if isinstance(metrics_list, list) else 'N/A'}")

if isinstance(metrics_list, list) and metrics_list:
    print(f"\nProcessing {len(metrics_list)} target languages...")

    for i, metrics_entry in enumerate(metrics_list):
        target_lang = metrics_entry.get('targetLanguage', 'unknown')
        print(f"\n--- Target Language {i+1}: {target_lang} ---")

        core_metrics = metrics_entry.get('coreMetrics', {})
        total_words = core_metrics.get('totalWords', 0)
        print(f"Total words: {total_words}")

        metrics_progress = metrics_entry.get('metricsProgress', {})
        print(f"Workflow steps: {list(metrics_progress.keys())}")

        for step_name, step_metrics in metrics_progress.items():
            print(f"  {step_name}:")
            print(f"    - Total words: {step_metrics.get('totalWordCount', 0)}")
            print(f"    - Words done: {step_metrics.get('wordsDone', 0)}")
            print(f"    - Words to do: {step_metrics.get('wordsToBeDone', 0)}")
else:
    print("No metrics data!")
    print(f"Metrics value: {metrics_list}")
