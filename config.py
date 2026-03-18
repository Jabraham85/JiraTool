"""
Avalanche Jira Template Creator — shared constants and configuration.
"""

APP_VERSION = "1.4.4"

# Update channels — each points to its own version manifest on GitHub
GITHUB_VERSION_URLS = {
    "stable": "https://raw.githubusercontent.com/Jabraham85/JiraTool/main/version.json",
    "experimental": "https://raw.githubusercontent.com/Jabraham85/JiraTool/main/version_experimental.json",
}
GITHUB_VERSION_URL = GITHUB_VERSION_URLS["stable"]  # legacy fallback

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
    # Sprint, Fix Version (Milestone), Time estimates
    "customfield_10020", "fixVersions",
    "timeoriginalestimate", "timeestimate",
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
    # Sprint, Milestone / Fix Version, Time estimates
    "Sprint", "Fix Version", "Original Estimate", "Remaining Estimate",
]

FETCHABLE_OPTION_FIELDS = {
    "Project key", "Issue Type", "Status", "Priority",
    "Components", "Labels", "Reporter",
    "Sprint", "Fix Version",
}

MULTISELECT_FIELDS = {"Labels"}

DEFAULT_KANBAN_COLUMNS = [
    {"name": "To Do",       "statuses": ["To Do", "Open", "Backlog", "New"]},
    {"name": "In Progress", "statuses": ["In Progress", "In Review", "In Development"]},
    {"name": "Done",        "statuses": ["Done", "Closed", "Resolved", "Complete", "Completed"]},
]

_JIRA_FIELDS_FALLBACK = [
    ("summary", "Summary"), ("description", "Description"),
    ("status", "Status"), ("assignee", "Assignee"),
    ("priority", "Priority"), ("labels", "Labels"),
    ("components", "Components"), ("issuetype", "Issue Type"),
    ("resolution", "Resolution"), ("reporter", "Reporter"),
    ("created", "Created"), ("updated", "Updated"),
]
