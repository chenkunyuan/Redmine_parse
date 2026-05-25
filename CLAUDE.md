# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Run

```bash
pip install -e .          # Install in dev mode (installs `redmine` CLI command)
redmine --help            # Verify installation
```

No test suite exists. No linter is configured.

## Architecture

**Two-layer design:** CLI layer (`cli.py`) → Client layer (`client.py`).

- `redmine_parse/cli.py` — Click-based CLI. Top-level `main` group parses `--url`/`--key` flags into Click context. Two subcommand groups: `projects` (only `list`) and `issues` (`list`, `show`, `create`, `update`). A shared `_get_client()` helper resolves URL/key from flags or env vars (`REDMINE_URL`, `REDMINE_API_KEY`), falling back to `https://redmine.sercomm.co.jp` as default URL. All commands output JSON to stdout via `_output()`.

- `redmine_parse/client.py` — `RedmineClient` wraps `python-redmine`'s `Redmine` class. All Redmine API exceptions (`AuthError`, `ForbiddenError`, `ResourceNotFoundError`, `ValidationError`, `ServerError`) are caught and re-raised as `RedmineClientError` with user-facing messages. The `_serialize()` helper converts python-redmine Resource objects to plain dicts by accessing `_decoded_attrs` (the library's internal raw API response) and recursing through dicts/lists. Datetime objects are ISO-formatted.

- `redmine_parse/__init__.py` — Exports only `RedmineClient`.

**Standalone scripts** (not part of the installed package):
- `export_hydrant_issues.py` — Exports issues from project 141 to `HYDRANT_Field_Issues.xlsx` with keyword-based auto-categorization (WIFI, WAN/Connection, Power/Reboot, etc.). Defaults to incremental mode (reads existing sheet, appends only new issue IDs). `--full` rebuilds from scratch.
- `generate_user_guide.py` — Generates `Redmine_Parse_User_Guide.docx` using `python-docx`.
