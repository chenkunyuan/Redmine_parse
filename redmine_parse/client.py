import datetime

from redminelib import Redmine
from redminelib.exceptions import (
    AuthError,
    ForbiddenError,
    ResourceNotFoundError,
    ValidationError,
    ServerError,
)


class RedmineClientError(Exception):
    pass


def _serialize(obj):
    """Convert python-redmine Resource objects to JSON-serializable dicts."""
    if isinstance(obj, (str, int, float, bool, type(None))):
        return obj

    if isinstance(obj, (datetime.datetime, datetime.date)):
        return obj.isoformat()

    # python-redmine Resource: use raw decoded API data
    if hasattr(obj, "_decoded_attrs"):
        return {k: _serialize(v) for k, v in obj._decoded_attrs.items()}

    if isinstance(obj, dict):
        return {k: _serialize(v) for k, v in obj.items()}

    if isinstance(obj, (list, tuple)):
        return [_serialize(item) for item in obj]

    return str(obj)


class RedmineClient:
    def __init__(self, url: str, api_key: str):
        self.url = url.rstrip("/")
        self.api_key = api_key
        try:
            self._redmine = Redmine(self.url, key=self.api_key)
        except Exception as e:
            raise RedmineClientError(f"Failed to initialize Redmine client: {e}") from e

    def get_projects(self):
        try:
            projects = self._redmine.project.all()
        except AuthError as e:
            raise RedmineClientError(f"Authentication failed: {e}") from e
        except ServerError as e:
            raise RedmineClientError(f"Server error: {e}") from e
        except Exception as e:
            raise RedmineClientError(f"Failed to fetch projects: {e}") from e
        return [_serialize(p) for p in projects]

    def get_issues(self, project_id: int, **filters):
        try:
            issues = self._redmine.issue.filter(project_id=project_id, **filters)
        except ResourceNotFoundError:
            raise RedmineClientError(f"Project with id={project_id} not found")
        except ForbiddenError:
            raise RedmineClientError(
                f"Access to project id={project_id} is forbidden"
            )
        except AuthError as e:
            raise RedmineClientError(f"Authentication failed: {e}") from e
        except Exception as e:
            raise RedmineClientError(f"Failed to fetch issues: {e}") from e
        return [_serialize(i) for i in issues]

    def get_issue(self, issue_id: int):
        try:
            issue = self._redmine.issue.get(issue_id)
        except ResourceNotFoundError:
            raise RedmineClientError(f"Issue #{issue_id} not found")
        except AuthError as e:
            raise RedmineClientError(f"Authentication failed: {e}") from e
        except ForbiddenError:
            raise RedmineClientError(f"Access to issue #{issue_id} is forbidden")
        except Exception as e:
            raise RedmineClientError(f"Failed to fetch issue #{issue_id}: {e}") from e
        return _serialize(issue)

    def create_issue(self, project_id: int, subject: str, **kwargs):
        try:
            issue = self._redmine.issue.create(
                project_id=project_id, subject=subject, **kwargs
            )
        except ResourceNotFoundError:
            raise RedmineClientError(f"Project with id={project_id} not found")
        except ValidationError as e:
            raise RedmineClientError(f"Validation error: {e}") from e
        except AuthError as e:
            raise RedmineClientError(f"Authentication failed: {e}") from e
        except Exception as e:
            raise RedmineClientError(f"Failed to create issue: {e}") from e
        return _serialize(issue)

    def update_issue(self, issue_id: int, **kwargs):
        try:
            issue = self._redmine.issue.get(issue_id)
        except ResourceNotFoundError:
            raise RedmineClientError(f"Issue #{issue_id} not found")
        except Exception as e:
            raise RedmineClientError(f"Failed to fetch issue #{issue_id}: {e}") from e

        notes = kwargs.pop("notes", None)
        try:
            if notes is not None:
                issue.notes = notes
            for key, value in kwargs.items():
                setattr(issue, key, value)
            issue.save()
        except ValidationError as e:
            raise RedmineClientError(f"Validation error: {e}") from e
        except AuthError as e:
            raise RedmineClientError(f"Authentication failed: {e}") from e
        except Exception as e:
            raise RedmineClientError(f"Failed to update issue #{issue_id}: {e}") from e

        return _serialize(self._redmine.issue.get(issue_id))
