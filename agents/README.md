# Scheduled tasks (Full setup)

Two scheduled tasks run the background work. Claude creates both during the Full setup from the
templates in this folder, after filling in the placeholders:

| Placeholder | Filled with |
|---|---|
| `{{DASHBOARD_URL}}` | the address of the user's own dashboard (created during setup) |
| `{{PROJECT_DIR}}` | the full Windows path of the project folder, e.g. `C:\Users\you\Documents\job-search-copilot` |
| `{{TIMEZONE}}` | the user's time zone, e.g. `America/Toronto` |

- `morning_agent.md` - runs every weekday morning (after the night scan): reads job-alert emails,
  scores the new jobs, spots employer replies in Gmail, refreshes the dashboard and writes the
  morning summary.
- `sync_agent.md` - runs when the user presses a button on the dashboard (it is booked by the page
  for the next minute): applies board moves, prepares documents, builds interview kits, refreshes
  the board.

Both tasks need the Claude desktop app on the computer that holds the project folder, because
they read and write files there. The user approves each task once when Claude creates it.
