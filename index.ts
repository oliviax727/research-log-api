/// <reference types="office-js" />

namespace ResearchLog {
  // User-defined types are not allowed so functionstring is explicitly defined
  export function runFunction(
    workbook: Excel.Workbook,
    functionString: functionStringOptions,
  ): void {
    // Get export function string from options
    const callbackFunction = functionMap.get(functionString);

    // Type safety
    if (typeof callbackFunction === "undefined") {
      return;
    }

    // Run called export function
    suspendSheetAndExecute(workbook, callbackFunction);

    return;
  }

  // Define a type based on the allowed export function inputs
  type functionStringOptions =
    | "Add Entry"
    | "Shift Entry"
    | "Pop Entry"
    | "Clock In"
    | "Clock Out";

  // Define a map of available functions
  const functionMap: Map<
    functionStringOptions,
    (appData: ExcelAppData) => void
  > = new Map([
    ["Add Entry", addEntry],
    ["Shift Entry", shiftEntry],
    ["Pop Entry", popEntry],
    ["Clock In", clockIn],
    ["Clock Out", clockOut],
  ]);

  // Container interface for arguments to entry functions
  interface ExcelAppData {
    workbook: Excel.Workbook;
    sheet: Excel.Worksheet;
    application: Excel.Application;
  }

  // Suspend calculations and run the given input export function
  export function suspendSheetAndExecute(
    workbook: Excel.Workbook,
    callback: (appData: ExcelAppData) => void,
  ): void {
    // Specify workbook and application
    const application = workbook.application;
    let sheet = workbook.worksheets.getItem("Logbook");

    // Turn off automatic calculations during the script.
    application.calculationMode = Excel.CalculationMode.manual;

    try {
      // Call and execute the callback
      callback({
        workbook: workbook,
        sheet: sheet,
        application: application,
      });
    } catch (error: unknown) {
      // Catch and check for errors, log if it is an error
      if (error instanceof Error) {
        console.log(error.message);
      } else {
        throw error;
      }
    } finally {
      // Manually calculate the file
      application.calculate(Excel.CalculationType.recalculate);
    }

    // Resume automatic calculations after the script finishes
    application.calculationMode = Excel.CalculationMode.automatic;
  }

  // Shift all entries down by one i.e. create a blank entry
  export function shiftEntry(appData: ExcelAppData): void {
    // Shift all rows down
    appData.sheet.getRange("B5:G5").insert(Excel.InsertShiftDirection.down);

    // Copy formatting from row below
    appData.sheet
      .getRange("B5")
      .copyFrom(
        appData.sheet.getRange("B6:G6"),
        Excel.RangeCopyType.all,
        false,
        false,
      );

    // Clear top row
    appData.sheet.getRange("B5").clear(Excel.ClearApplyTo.contents);
    appData.sheet.getRange("F5:G5").clear(Excel.ClearApplyTo.contents);
  }

  // Add a new entry
  export function addEntry(appData: ExcelAppData): void {
    // Shift Entry
    shiftEntry(appData);

    // Get current date
    let date = new Date(Date.now());

    // Get custom/override entry cells
    let customDescriptionEntry = appData.sheet.getRange("J31");
    let customDateEntry = appData.sheet.getRange("L32");
    let customHoursEntry = appData.sheet.getRange("N32");

    // Copy entry description from either defaults of override
    if (customDescriptionEntry.values?.[0]?.[0]) {
      appData.sheet
        .getRange("F5")
        .copyFrom(
          customDescriptionEntry,
          Excel.RangeCopyType.values,
          false,
          false,
        );
    } else {
      appData.sheet.getRange("F5").values = [[
        "Unspecified work on " + date.toLocaleDateString(),
      ]];
    }

    // Copy entry hours from either defaults of override
    if (customHoursEntry.values?.[0]?.[0]) {
      appData.sheet
        .getRange("G5")
        .copyFrom(customHoursEntry, Excel.RangeCopyType.values, false, false);
    } else {
      appData.sheet
        .getRange("G5")
        .copyFrom(
          appData.sheet.getRange("L22"),
          Excel.RangeCopyType.values,
          false,
          false,
        );
    }

    // Copy entry date from either defaults of override
    if (customDateEntry.values?.[0]?.[0]) {
      appData.sheet
        .getRange("B5")
        .copyFrom(customDateEntry, Excel.RangeCopyType.values, false, false);
    } else {
      appData.sheet
        .getRange("B5")
        .copyFrom(
          appData.sheet.getRange("T5"),
          Excel.RangeCopyType.values,
          false,
          false,
        );
    }

    // Clear custom entries
    appData.sheet.getRange("J31:N31").clear(Excel.ClearApplyTo.contents);
    appData.sheet.getRange("L32").clear(Excel.ClearApplyTo.contents);
    appData.sheet.getRange("N32").clear(Excel.ClearApplyTo.contents);
  }

  // Remove the most recent entry
  export function popEntry(appData: ExcelAppData) {
    // Delete top row and shift up
    appData.sheet.getRange("B5:G5").delete(Excel.DeleteShiftDirection.up);
  }

  // Set the start time for the day
  export function clockIn(appData: ExcelAppData) {
    // Get current time and round to nearest 30 mins
    let date = new Date(Date.now());
    let time = roundToNearest30(date);

    // Set start time cell to given time
    let timeRange = appData.sheet.getRange("L18");
    timeRange.values = [[time.toLocaleTimeString()]];
  }

  // Round a date to the nearest 30 minutes
  export function roundToNearest30(date: Date) {
    const minutes = 30;
    const ms = 1000 * 60 * minutes;

    return new Date(Math.floor(date.getTime() / ms) * ms);
  }

  // Add another entry and clear the start time for the day
  export function clockOut(appData: ExcelAppData) {
    // Add Entry
    addEntry(appData);

    // Clear start time
    appData.sheet.getRange("L18").clear(Excel.ClearApplyTo.contents);
  }
}

export { ResearchLog };
