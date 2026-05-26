"use strict";
/// <reference types="office-js" />
var ResearchLog;
(function (ResearchLog) {
    // User-defined types are not allowed so functionstring is explicitly defined
    function runFunction(workbook, functionString) {
        // Get function string from options
        const callbackFunction = functionMap.get(functionString);
        // Type safety
        if (typeof callbackFunction === "undefined") {
            return;
        }
        // Run called function
        suspendSheetAndExecute(workbook, callbackFunction);
        return;
    }
    // Define a map of available functions
    const functionMap = new Map([
        ["Add Entry", addEntry],
        ["Shift Entry", shiftEntry],
        ["Pop Entry", popEntry],
        ["Clock In", clockIn],
        ["Clock Out", clockOut],
    ]);
    // Suspend calculations and run the given input function
    function suspendSheetAndExecute(workbook, callback) {
        // Specify workbook and application
        const application = workbook.application;
        let sheet = workbook.worksheets.getItem("Logbook");
        // Turn off automatic calculations during the script.
        application.calculationMode = Excel.CalculationMode.manual;
        try {
            // Call and execute the callback function
            callback({
                workbook: workbook,
                sheet: sheet,
                application: application,
            });
        }
        catch (error) {
            // Catch and check for errors, log if it is an error
            if (error instanceof Error) {
                console.log(error.message);
            }
            else {
                throw error;
            }
        }
        finally {
            // Manually calculate the file
            application.calculate(Excel.CalculationType.recalculate);
        }
        // Resume automatic calculations after the script finishes
        application.calculationMode = Excel.CalculationMode.automatic;
    }
    // Shift all entries down by one i.e. create a blank entry
    function shiftEntry(appData) {
        // Shift all rows down
        appData.sheet.getRange("B5:G5").insert(Excel.InsertShiftDirection.down);
        // Copy formatting from row below
        appData.sheet
            .getRange("B5")
            .copyFrom(appData.sheet.getRange("B6:G6"), Excel.RangeCopyType.all, false, false);
        // Clear top row
        appData.sheet.getRange("B5").clear(Excel.ClearApplyTo.contents);
        appData.sheet.getRange("F5:G5").clear(Excel.ClearApplyTo.contents);
    }
    // Add a new entry
    function addEntry(appData) {
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
                .copyFrom(customDescriptionEntry, Excel.RangeCopyType.values, false, false);
        }
        else {
            appData.sheet.getRange("F5").values = [[
                    "Unspecified work on " + date.toLocaleDateString(),
                ]];
        }
        // Copy entry hours from either defaults of override
        if (customHoursEntry.values?.[0]?.[0]) {
            appData.sheet
                .getRange("G5")
                .copyFrom(customHoursEntry, Excel.RangeCopyType.values, false, false);
        }
        else {
            appData.sheet
                .getRange("G5")
                .copyFrom(appData.sheet.getRange("L22"), Excel.RangeCopyType.values, false, false);
        }
        // Copy entry date from either defaults of override
        if (customDateEntry.values?.[0]?.[0]) {
            appData.sheet
                .getRange("B5")
                .copyFrom(customDateEntry, Excel.RangeCopyType.values, false, false);
        }
        else {
            appData.sheet
                .getRange("B5")
                .copyFrom(appData.sheet.getRange("T5"), Excel.RangeCopyType.values, false, false);
        }
        // Clear custom entries
        appData.sheet.getRange("J31:N31").clear(Excel.ClearApplyTo.contents);
        appData.sheet.getRange("L32").clear(Excel.ClearApplyTo.contents);
        appData.sheet.getRange("N32").clear(Excel.ClearApplyTo.contents);
    }
    // Remove the most recent entry
    function popEntry(appData) {
        // Delete top row and shift up
        appData.sheet.getRange("B5:G5").delete(Excel.DeleteShiftDirection.up);
    }
    // Set the start time for the day
    function clockIn(appData) {
        // Get current time and round to nearest 30 mins
        let date = new Date(Date.now());
        let time = roundToNearest30(date);
        // Set start time cell to given time
        let timeRange = appData.sheet.getRange("L18");
        timeRange.values = [[time.toLocaleTimeString()]];
    }
    // Round a date to the nearest 30 minutes
    function roundToNearest30(date) {
        const minutes = 30;
        const ms = 1000 * 60 * minutes;
        return new Date(Math.floor(date.getTime() / ms) * ms);
    }
    // Add another entry and clear the start time for the day
    function clockOut(appData) {
        // Add Entry
        addEntry(appData);
        // Clear start time
        appData.sheet.getRange("L18").clear(Excel.ClearApplyTo.contents);
    }
})(ResearchLog || (ResearchLog = {}));
export { ResearchLog };
//# sourceMappingURL=index.js.map