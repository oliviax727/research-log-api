# research-log-api

This repository contains a portable core module plus environment adapters for Office Scripts (ExcelScript) and Office Add-ins (office-js).

Quick summary

- `src/core.ts` — pure logic with no host-specific types.
- `src/adapters/office-script-adapter.ts` — small adapter to use in Office Scripts.
- `src/adapters/office-js-adapter.ts` — small adapter to use in an Office Add-in.

Build

Install dev deps and run the build scripts:

```bash
npm install
npm run build:bundle    # produces dist/office-js-adapter.js (ESM)
npm run build:script    # concatenates core + office-script adapter into dist/office-script-single.ts
```

Office Scripts

- Open `dist/office-script-single.ts`, paste into the Office Scripts editor, and call `runFunction(workbook, "Add Entry")` (adapt as needed).

Office Add-ins

- Host `dist/office-js-adapter.js` on your add-in server and import it from your add-in code.

Deno / Remote imports

Deno can import `index.ts` directly from `raw.githubusercontent.com`. Ensure the exported API is ESM-compatible when importing into other runtimes.
