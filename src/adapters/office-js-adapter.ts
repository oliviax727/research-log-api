// Adapter for Office Add-ins (office-js)
// Maps `Excel.Workbook` to portable core APIs.

import type { Entry, FunctionStringOptions } from "../core";
import { addEntry, popEntry, shiftEntry, clockIn, clockOut } from "../core";

export async function runFunction(workbook: any, functionString: FunctionStringOptions): Promise<void> {
  const sheet = workbook.worksheets.getItem("Logbook");

  // TODO: implement mapping from worksheet ranges to `Entry[]` using office-js APIs
  const entries: Entry[] = [];

  let result: Entry[] = entries;

  switch (functionString) {
    case "Add Entry":
      result = addEntry(entries, { date: new Date().toISOString(), notes: "Added via Add-in" });
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

  // TODO: write `result` back into the workbook ranges via office-js APIs
}
