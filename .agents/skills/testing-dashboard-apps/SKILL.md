---
name: testing-dashboard-apps
description: Test PaperHub and ProjectHub dashboard apps end-to-end. Use when verifying UI terminology, export functionality, or pane content changes.
---

# Testing PaperHub / ProjectHub Dashboard Apps

## Quick Start

1. Navigate to the app directory (`paper-dashboard/` or `project-hub/`)
2. Run `npm install && npm run dev` to start the Vite dev server at `http://localhost:5173`
3. Open the browser and navigate to `http://localhost:5173`

## Key Verification Points

### Terminology Check (9 Panes)
The dashboard has 9 configurable panes. When verifying terminology changes:
- Read the full HTML output (it gets truncated — check the overflow file)
- Key panes to check: Pane 1 (Title/Team), Pane 7 (Metrics), Pane 8 (Budget/Costs), Pane 9 (Team Tasks)
- For PaperHub: Authors, Journal, IF, APC, Co-author Tasks, Word Count
- For ProjectHub: Team Members, Client, Priority, Budget, Team Tasks, Sprints

### Add Dialog
- Click "+ Add New [Project/Paper]" button in sidebar to open the dialog
- The submit button text is at the bottom of the dialog — scroll down or check the HTML overflow file
- The button is `disabled` until a title is typed

### Export Testing
- **JSON Export**: Click "Export JSON" — downloads as `projects.json` or `papers.json`
- **YAML Export**: Click "Export YAML" — downloads as `projects.yaml` or `papers.yaml`
- Check the first line of YAML for the root key (`projects:` vs `papers:`)
- Downloaded files appear at `/tmp/chisel_browser_downloads/`

### LocalStorage Keys
- The `useLocalStorage` hook only writes to localStorage when `setValue` is called (not on initial render)
- On a fresh load, localStorage will be empty — the app uses in-memory defaults from `samplePapers`
- To verify key names, check the source code in `src/context/DashboardContext.tsx` for `useLocalStorage` calls
- PaperHub uses `paperhub-*` prefix, ProjectHub uses `projecthub-*` prefix

### Tooltip Verification
- Tooltips are set via `title` attribute on HTML elements
- Read the HTML output to verify tooltip text rather than trying to hover
- Budget/APC tooltip is on the clickable div in the Budget & Funding pane

## Common Issues

- HTML output from the browser tool gets truncated at ~150 lines. Always read the overflow file at `/tmp/devin-remote-overflows-1000/*/content.txt` to see the full page content.
- The YAML serializer is custom (not a library) — check `DashboardContext.tsx` for the implementation.
- YAML import requires file dialog interaction which is hard to test at runtime. Verify backward compatibility logic in source code instead.

## Build Verification

```bash
cd <app-dir>
npm run lint
npm run build
```

Both must pass with no errors.

## Devin Secrets Needed

No secrets required — both apps run fully locally with sample data and localStorage persistence.
