# Supabase Collaboration Plan

## Goal

Turn the dashboard into a shared online planning board where multiple people can edit the same sprint at the same time and see updates appear live.

This is a Google-Sheets-like collaboration layer for the board, not a full spreadsheet engine. The source of truth becomes Supabase tables instead of a local CSV file.

## Why Supabase

- Works well with structured task data
- Gives us Postgres plus Realtime subscriptions
- Fits the current static HTML app
- Easier to reason about than building full CRDT spreadsheet syncing
- Still supports CSV import and CSV export

## Collaboration Model

### Workspace

A `workspace` is one shared planning board.

Example:

- `sprint-1`
- `sprint-2-rehearsal`
- `may-launch-plan`

Each workspace stores:

- the original CSV rows used for export
- the current task state used for live collaboration

### Tasks

Each task is stored as one row in `planning_tasks`.

Live edits update:

- `assignee`
- `start_date`
- `end_date`

The app recalculates working days in the browser from the date range.

## User Flow

### First publish

1. Open the board
2. Upload a CSV
3. Enter Supabase URL, anon key, and workspace ID
4. Connect
5. Publish the current board to the workspace

### Teammate join

1. Open the same app
2. Enter the same Supabase URL, anon key, and workspace ID
3. Connect
4. Load the shared workspace

### Live editing

- Drag task bars to move dates
- Resize task bars to change duration
- Click tasks to open the overlay editor
- Change assignee in the overlay
- Changes are saved to Supabase
- Other connected clients receive the update through Realtime

## Current v1 Behavior

- Last write wins
- Realtime sync is task-based
- CSV export stays available as long as the workspace was published from a CSV
- No user auth or edit ownership yet

## Suggested Future Improvements

### Presence

Show:

- who is online
- who is editing which task
- short-lived edit locks while dragging

This can be added later with Liveblocks or a simple Supabase presence channel.

### Auth

Move from anonymous access to:

- Supabase Auth with email login, or
- office SSO if needed

### History

Add:

- audit trail
- undo history
- last editor per task

### Safer publishing

Add:

- publish confirmation
- workspace versioning
- snapshot restore

## App Setup Notes

The static app expects these values at runtime:

- Supabase project URL
- Supabase anon key
- workspace ID

The current implementation stores these in browser `localStorage` for convenience.

## Security Note

The v1 schema below is intentionally simple so the team can move quickly.

If the repo is public or the tool will be used outside a trusted office environment, tighten it with:

- authenticated users
- RLS policies scoped by user or team
- server-validated workspace membership

## Required Supabase Tables

See:

- [supabase/schema.sql](../supabase/schema.sql)

## Deployment Notes

Because the app is static, it can be deployed with:

- GitHub Pages
- Netlify
- Vercel

Supabase handles the shared online state separately.
