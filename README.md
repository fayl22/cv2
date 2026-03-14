# SharePoint Issue Tracker

A lightweight issue tracker app that uses a SharePoint list as the backend through the SharePoint REST API.

## Files

- `issue-tracker.html` — UI for managing issue tracker settings and items.
- `sharepoint-issue-tracker.js` — SharePoint REST integration and issue CRUD logic.

## SharePoint List Setup

Create a SharePoint list (default name: `Issues`) with at least these columns:

- **Title** (single line of text)
- **Description** (multiple lines of text)
- **Priority** (choice: Low, Medium, High, Critical)
- **IssueStatus** (choice: Open, In Progress, Blocked, Resolved, Closed)

## Usage

1. Open `issue-tracker.html` in a browser.
2. Enter:
   - SharePoint site URL (example: `https://contoso.sharepoint.com/sites/Engineering`)
   - List title (`Issues` by default)
   - OAuth Bearer token (optional, if required by your auth flow)
3. Click **Save Config**.
4. Click **Load Issues** to fetch current issues.
5. Use the create form to add new issues.
6. Use per-row actions to toggle status or delete issues.

## Notes

- For write operations, this app requests a SharePoint form digest from `/_api/contextinfo`.
- If you run this app outside SharePoint-hosted pages, ensure CORS/authentication allows access.
