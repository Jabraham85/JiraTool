# Avalanche Jira Template Creator — Voice Actor Walkthrough Script

**Purpose:** Narration script for a video walkthrough covering all features and settings.  
**Tone:** Clear, professional, friendly. Pause briefly between sections.

---

## INTRO

Welcome to Avalanche Jira Template Creator — a desktop app for creating, managing, and uploading Jira tickets in bulk. The sidebar has Templates, Jira, Bundle, and Help. The main area shows ticket tabs or a list view. Let's go through each section.

---

## TEMPLATES SECTION

The **Templates** list shows your saved templates. Click to select, double-click to open a new ticket. **New**, **Duplicate**, and **Delete** manage templates. **Save Template** saves the active ticket as a blueprint; **Save All** saves every open tab with changes.

**Import Template from CSV Row(s)** creates templates from a CSV file. **Import CSV rows into current tab** pastes CSV data into the active ticket. **Remove Field From Template** strips a field from the selected template. **Check All**, **Uncheck All**, and **Collapse Unincluded Fields** bulk-toggle which fields are visible and included in uploads.

---

## JIRA SECTION

**Set Jira API** stores your base URL, email, and API token — generate one at id.atlassian.com. **Test Connection** verifies credentials. **Fetch My Issues** pulls tickets from Jira with filters: scope, project, labels, components, and more. **Auto-Fetch Settings** runs that fetch automatically on startup. **Refresh Fetched Tickets** re-fetches all previously downloaded tickets.

**Configure Reminders** lets you set internal priority levels, enable stale-ticket warnings when a ticket hasn't been updated in X days, choose which fields to ignore for stale detection, and toggle the tutorial on first startup.

**Export Bundle** saves your bundle to a `.avl` file to share; **Import Bundle** loads one.

---

## HELP SECTION

**Show Tutorial** runs the interactive guided tour — connect to Jira, create tickets, save templates, use variables, fetch, bulk import, and upload. Available anytime from Help.

---

## BUNDLE SECTION

A bundle is a set of tickets you upload to Jira together. **Add** puts the active tab in the bundle. **Remove** and **Clear** manage the list; **Rename** gives it a custom name. **Jira** opens the selected item in your browser. **Export** saves to `.avl` — same as in the Jira section. **Upload Bundle to Jira** sends every ticket in one go, creating or updating issues as needed.

**Bulk Import** creates many tickets from a template. Choose a template, paste text, and click Import — each block becomes a ticket and is added to the bundle.

**Setup**: Blank lines separate blocks. The first line of each block is the ticket name. The lines under it are details. Put exclamation marks in the template where you want those details to go. Each detail line replaces an exclamation in order — first line goes in the first exclamation, second line in the second, and so on. If there are more detail lines than exclamation marks, all remaining lines go into the last exclamation. If there are no exclamation marks, the details are appended to the end of the ticket.

**Parse mode**: Structured uses this setup. Replace, Prepend, and Append modes use each line as a summary instead.

**Options**: line delimiter — newline, comma, semicolon, tab, or pipe; exclude empty lines or lines starting with a prefix like hash or slash; separator for prepend and append; content format — smart list, paragraphs, bullet list, or ordered list — for how detail lines are inserted. If your text has blank lines, structured parsing runs automatically.

---

## TOOLBAR

Above the tabs: **Search fields** filters the form. **New Tab**, **Close Tab**, **Duplicate Tab**, and **Save** manage tabs. **Toggle List View** switches between tabs and a sortable table. **Variable** opens the variable dialog — select text, right-click Define Variable, use Insert Variable anywhere; they persist in templates and resolve on upload. **Refresh from Jira** pulls the latest data for the active ticket.

---

## LIST VIEW

In list view you see a table of all tickets. **Search list**, **Scope**, and **Folder** filters narrow it; **Manage Folders** organizes folders. **Import CSV into list** adds tickets from a file. **Move to Folder**, **Mass Edit**, **Open selected as Tabs**, **Add selected to Bundle**, **Remove Selected**, and **Clear List** work on checked or selected rows. **Jira** opens the selected ticket in browser; **Open All as Tabs** opens every visible ticket.

---

## TICKET FORM

Each tab shows a form. The **Internal** dropdown is a local priority that never uploads. **Summary** and **Description** — with rich ADF via the Edit button and Jira Preview. Standard Jira fields: Issue Type, Status, Priority, Labels, Components, Assignee — use the refresh icon to reload options from Jira. **Epic** links to an Epic via Search. **Issue Links** adds blocks, relates-to, and similar. **Comments** shows the thread; **Add Comment** posts. **Attachment** holds semicolon-separated file paths; the app uploads them when you upload the ticket. Right-click in any text field for Define Variable and Insert Variable.

---

## WELCOME TAB

The landing page shows **New / Updated Tickets** and **High Internal Priority** — double-click to open. Buttons: Fetch My Issues, Refresh All Tickets, Jira. Sync status at the top shows when data was last refreshed.

---

## OUTRO

That covers Avalanche Jira Template Creator — connect to Jira, create templates, use variables, bulk import, manage bundles, and upload in bulk. For a hands-on tour, click **Show Tutorial** in Help. Thanks for watching.
