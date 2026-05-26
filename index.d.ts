declare namespace ResearchLog {
    export function runFunction(workbook: Excel.Workbook, functionString: functionStringOptions): void;
    type functionStringOptions = "Add Entry" | "Shift Entry" | "Pop Entry" | "Clock In" | "Clock Out";
    interface ExcelAppData {
        workbook: Excel.Workbook;
        sheet: Excel.Worksheet;
        application: Excel.Application;
    }
    export function suspendSheetAndExecute(workbook: Excel.Workbook, callback: (appData: ExcelAppData) => void): void;
    export function shiftEntry(appData: ExcelAppData): void;
    export function addEntry(appData: ExcelAppData): void;
    export function popEntry(appData: ExcelAppData): void;
    export function clockIn(appData: ExcelAppData): void;
    export function roundToNearest30(date: Date): Date;
    export function clockOut(appData: ExcelAppData): void;
    export {};
}
export { ResearchLog };
//# sourceMappingURL=index.d.ts.map