---
name: testing-paperhub
description: Test PaperHub paper management dashboard end-to-end. Use when verifying inline editing, Gantt chart, or pane UI changes.
---

# Testing PaperHub

## Setup

```bash
cd paper-dashboard && npm install && npm run dev -- --port 5173
```

App runs at `http://localhost:5173`. If port is taken, Vite auto-increments.

## Key Test Areas

### Gantt Chart (GanttChart.tsx)
- Events **without** `endDate` should extend bars to Today (ongoing)
- Events **with** `endDate` (even if same as `startDate`) render at that date (completed/point marker)
- Verify bar widths via console: `document.querySelectorAll('.gantt-bar')` or check `style.width` on bar elements
- Test with PMEA paper (mix of ongoing + completed events) and Cryoanesthesia (completed single-day events)

### Inline Editing (All 9 Panes)
- Click text → input/textarea appears → type → Enter saves, Escape cancels
- Hover shows dashed underline on editable fields
- Test multi-character input (not just single char) to catch remount bugs
- **Important pattern**: Never define sub-components (like `StatCard`) inside a parent component body. This creates new function references on each render, causing React to unmount/remount, which fires `onBlur` mid-edit and breaks multi-character input. Always extract to module scope.

### Persistence
- Data persists in localStorage
- Clear with `localStorage.clear()` before fresh test runs
- Edits should survive paper switching (select different paper → select back)

## Browser Automation Workarounds

### Sidebar Paper Selection
The sidebar items might not respond to direct `devinid` clicks. Use React props workaround:
```javascript
const el = document.querySelector('[devinid="X"]');
const reactProps = Object.keys(el).find(k => k.startsWith('__reactProps'));
if (reactProps) el[reactProps].onClick();
```

### Statistics Pane Location
Statistics pane (Pane 7) is below the fold. Scroll into view:
```javascript
document.querySelector('[devinid="104"]').scrollIntoView({ behavior: 'instant', block: 'center' });
```
Devinids may change; search for elements with `title="Click to edit count"` to find stat values.

### Number Input Editing
For `type="number"` inputs, use native setter to set values:
```javascript
const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
nativeInputValueSetter.call(input, '5000');
input.dispatchEvent(new Event('input', { bubbles: true }));
```

### Add Paper Dialog
The "+ New Paper" button might not respond to browser automation clicks. This is a known limitation — mark as untested if blocked.

## Sample Papers for Testing
- **Zero-cal PMEA**: Mix of ongoing (Revision, no endDate) and completed (Submission, endDate=startDate) events
- **GDP tempo-effect**: Ongoing submission (no endDate)
- **Cryoanesthesia review**: Completed single-day events (Accepted, Submission both have endDate=startDate)

## Devin Secrets Needed
None — PaperHub is a fully client-side app with no backend authentication.
