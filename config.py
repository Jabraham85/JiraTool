"""
Avalanche Jira Template Creator — shared constants and configuration.
"""

APP_VERSION = "1.0.6"
GITHUB_VERSION_URL = "https://raw.githubusercontent.com/Jabraham85/JiraTool/main/version.json"

TEMPLATES_FILE = "templates.json"
DEBUG_LOG = "jira_debug.log"

FETCH_FIELDS = [
    "summary", "description", "issuetype", "status", "project", "priority",
    "assignee", "reporter", "created", "updated", "labels", "components",
    "environment", "attachment", "comment",
    # Epic / parent hierarchy
    "parent", "customfield_10014", "customfield_10011",
    # Formal issue links (blocks / relates-to / etc.)
    "issuelinks",
]

HEADERS = [
    "Summary", "Description", "Description ADF", "Issue key", "Issue id",
    "Issue Type", "Status", "Project key", "Project name", "Components",
    "Labels", "Priority", "Assignee", "Reporter", "Creator", "Created",
    "Updated", "Environment", "Attachment", "Comment",
    # Epic relationship (stored separately from sub-task parent)
    "Epic Link", "Epic Name", "Epic Children",
    # Formal issue links (blocks / relates-to / etc.)
    "Issue Links",
    # Sub-task parent
    "Parent", "Parent key", "Parent summary",
    "Status Category", "Variables",
]

FETCHABLE_OPTION_FIELDS = {
    "Project key", "Issue Type", "Status", "Priority",
    "Components", "Labels", "Reporter",
}

MULTISELECT_FIELDS = {"Labels"}

_JIRA_FIELDS_FALLBACK = [
    ("summary", "Summary"), ("description", "Description"),
    ("status", "Status"), ("assignee", "Assignee"),
    ("priority", "Priority"), ("labels", "Labels"),
    ("components", "Components"), ("issuetype", "Issue Type"),
    ("resolution", "Resolution"), ("reporter", "Reporter"),
    ("created", "Created"), ("updated", "Updated"),
]
