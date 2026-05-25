import json
import os
import sys

import click

from .client import RedmineClient, RedmineClientError

DEFAULT_URL = "https://redmine.sercomm.co.jp"


def _get_client(url: str, key: str) -> RedmineClient:
    url = url or os.environ.get("REDMINE_URL", DEFAULT_URL)
    key = key or os.environ.get("REDMINE_API_KEY", "")

    if not key:
        raise click.UsageError(
            "API key is required. Set REDMINE_API_KEY environment variable "
            "or pass --key."
        )

    try:
        return RedmineClient(url, key)
    except RedmineClientError as e:
        raise click.ClickException(str(e))


def _output(data) -> None:
    text = json.dumps(data, indent=2, ensure_ascii=False, default=str)
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    click.echo(text)


@click.group()
@click.option("--url", default=None, help="Redmine URL (default: REDMINE_URL env or https://redmine.sercomm.co.jp)")
@click.option("--key", default=None, help="Redmine API key (default: REDMINE_API_KEY env)")
@click.pass_context
def main(ctx, url, key):
    ctx.ensure_object(dict)
    ctx.obj["url"] = url
    ctx.obj["key"] = key


@main.group()
def projects():
    """List and manage projects."""


@projects.command("list")
@click.pass_context
def projects_list(ctx):
    """List all projects."""
    client = _get_client(ctx.obj["url"], ctx.obj["key"])
    try:
        result = client.get_projects()
    except RedmineClientError as e:
        raise click.ClickException(str(e))
    _output(result)


@main.group()
def issues():
    """List, show, create, and update issues."""


@issues.command("list")
@click.option("--project-id", "-p", type=int, required=True, help="Project ID")
@click.pass_context
def issues_list(ctx, project_id):
    """List issues in a project."""
    client = _get_client(ctx.obj["url"], ctx.obj["key"])
    try:
        result = client.get_issues(project_id)
    except RedmineClientError as e:
        raise click.ClickException(str(e))
    _output(result)


@issues.command("show")
@click.argument("issue_id", type=int)
@click.pass_context
def issues_show(ctx, issue_id):
    """Show full details of a single issue."""
    client = _get_client(ctx.obj["url"], ctx.obj["key"])
    try:
        result = client.get_issue(issue_id)
    except RedmineClientError as e:
        raise click.ClickException(str(e))
    _output(result)


@issues.command("create")
@click.option("--project-id", "-p", type=int, required=True, help="Project ID")
@click.option("--subject", "-s", required=True, help="Issue subject")
@click.option("--description", "-d", default=None, help="Issue description")
@click.option("--tracker-id", type=int, default=None, help="Tracker ID")
@click.option("--status-id", type=int, default=None, help="Status ID")
@click.option("--priority-id", type=int, default=None, help="Priority ID")
@click.option("--assigned-to-id", type=int, default=None, help="Assignee user ID")
@click.pass_context
def issues_create(ctx, project_id, subject, description, tracker_id, status_id,
                  priority_id, assigned_to_id):
    """Create a new issue."""
    kwargs = {}
    if description is not None:
        kwargs["description"] = description
    if tracker_id is not None:
        kwargs["tracker_id"] = tracker_id
    if status_id is not None:
        kwargs["status_id"] = status_id
    if priority_id is not None:
        kwargs["priority_id"] = priority_id
    if assigned_to_id is not None:
        kwargs["assigned_to_id"] = assigned_to_id

    client = _get_client(ctx.obj["url"], ctx.obj["key"])
    try:
        result = client.create_issue(project_id, subject, **kwargs)
    except RedmineClientError as e:
        raise click.ClickException(str(e))
    _output(result)


@issues.command("update")
@click.argument("issue_id", type=int)
@click.option("--subject", "-s", default=None, help="New subject")
@click.option("--description", "-d", default=None, help="New description")
@click.option("--tracker-id", type=int, default=None, help="New tracker ID")
@click.option("--status-id", type=int, default=None, help="New status ID")
@click.option("--priority-id", type=int, default=None, help="New priority ID")
@click.option("--assigned-to-id", type=int, default=None, help="New assignee user ID")
@click.option("--notes", default=None, help="Add a note/comment")
@click.pass_context
def issues_update(ctx, issue_id, subject, description, tracker_id, status_id,
                  priority_id, assigned_to_id, notes):
    """Update an existing issue."""
    kwargs = {}
    if subject is not None:
        kwargs["subject"] = subject
    if description is not None:
        kwargs["description"] = description
    if tracker_id is not None:
        kwargs["tracker_id"] = tracker_id
    if status_id is not None:
        kwargs["status_id"] = status_id
    if priority_id is not None:
        kwargs["priority_id"] = priority_id
    if assigned_to_id is not None:
        kwargs["assigned_to_id"] = assigned_to_id
    if notes is not None:
        kwargs["notes"] = notes

    if not kwargs:
        raise click.UsageError("At least one field to update is required.")

    client = _get_client(ctx.obj["url"], ctx.obj["key"])
    try:
        result = client.update_issue(issue_id, **kwargs)
    except RedmineClientError as e:
        raise click.ClickException(str(e))
    _output(result)


if __name__ == "__main__":
    main()
