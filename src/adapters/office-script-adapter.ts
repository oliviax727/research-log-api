// Adapter for Office Scripts (Excel on the web)
// This module maps `ExcelScript.Workbook` to the portable core APIs.
// Paste the concatenated output (core + this adapter) into the Office Scripts editor.

import type { Entry, FunctionStringOptions } from "../core";
import { addEntry, popEntry, shiftEntry, clockIn, clockOut } from "../core";

// The following signature is intended for Office Scripts runtime:
export async function runFunction(workbook: any, functionString: FunctionStringOptions) {
  const sheet = workbook.getWorksheet("Logbook");

  // TODO: implement mapping from worksheet ranges to `Entry[]`.
  // Example placeholder loader — replace with real range reads.
  const entries: Entry[] = [];

  let result: Entry[] = entries;

  switch (functionString) {
    case "Add Entry":
      result = addEntry(entries, { date: new Date().toISOString(), notes: "Added via Office Script" });
      break;
    case "Shift Entry":
      result = shiftEntry(entries);
      break;
    case "Pop Entry":
      result = popEntry(entries);
      break;
    case "Clock In":
      result = clockIn(entries, new Date().toISOString());
      break;
    case "Clock Out":
      result = clockOut(entries, new Date().toISOString());
      break;
  }

  // TODO: write `result` back into the workbook ranges
  // e.g. sheet.getRange("A1").setValues(...)
}
