# research-log-api

This repository contains the `ResearchLog` namespace with reusable functions for Excel workbook operations, optimized for Office Scripts and Office Add-ins.

Quick summary

- `index.ts` — core ResearchLog namespace with functions for managing entries and tracking time.
- `package.json` — build scripts to compile for Office Scripts use.

Build

Install dev deps and compile:

```bash
npm install
npm run build              # standard TypeScript compilation
npm run build:office-script # compile index.ts for Office Scripts (outputs dist/office-script.ts)
```

Office Scripts

- Run `npm run build:office-script` to generate `dist/office-script.ts`.
- Open the generated file and paste the code into the Office Scripts editor.
- Call `ResearchLog.runFunction(workbook, "Add Entry")` from another script or button action.

Functions available

- `ResearchLog.addEntry(appData)` — add a new entry to the logbook.
- `ResearchLog.popEntry(appData)` — remove the most recent entry.
- `ResearchLog.shiftEntry(appData)` — create a blank entry at the top.
- `ResearchLog.clockIn(appData)` — set the start time (cell L18) to current rounded time.
- `ResearchLog.clockOut(appData)` — add an entry using the start time and clear L18.
- `ResearchLog.roundToNearest30(date)` — round a Date to the nearest 30 minutes.
