# Redmine Parse

CLI tool for interacting with Redmine via the REST API.

## Install

```
pip install -e .
```

## Setup

Set environment variables:

```bash
export REDMINE_URL="https://redmine.sercomm.co.jp"
export REDMINE_API_KEY="your-api-key"
```

Or pass `--url` and `--key` flags on every command.

## Usage

```bash
# List all projects
redmine projects list

# List issues in a project
redmine issues list --project-id 33

# Show full issue details
redmine issues show 3471

# Create an issue
redmine issues create --project-id 33 --subject "Bug report" --description "Details here"

# Update an issue
redmine issues update 3471 --notes "Added a note" --status-id 2
```

All commands output JSON to stdout.
