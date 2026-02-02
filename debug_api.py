#!/usr/bin/env python3
"""Debug script to see actual XTM API response structure"""

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
print("Fetching projects list...")
response = requests.get(f"{base_url}/projects", headers=headers)
projects = response.json()

if projects and len(projects) > 0:
    first_project = projects[0]
    project_id = first_project.get('id')

    print(f"\n=== SAMPLE PROJECT (ID: {project_id}) ===")
    print(json.dumps(first_project, indent=2))

    # Get metrics for this project
    print(f"\n=== METRICS FOR PROJECT {project_id} ===")
    try:
        metrics_response = requests.get(f"{base_url}/projects/{project_id}/metrics", headers=headers)
        if metrics_response.status_code == 200:
            metrics = metrics_response.json()
            print(json.dumps(metrics, indent=2))
        else:
            print(f"Status: {metrics_response.status_code}")
    except Exception as e:
        print(f"Error: {e}")

    # Try workflow endpoint
    print(f"\n=== WORKFLOW FOR PROJECT {project_id} ===")
    try:
        workflow_response = requests.get(f"{base_url}/projects/{project_id}/workflow", headers=headers)
        if workflow_response.status_code == 200:
            workflow = workflow_response.json()
            print(json.dumps(workflow, indent=2))
        else:
            print(f"Status: {workflow_response.status_code}")
    except Exception as e:
        print(f"Error: {e}")
