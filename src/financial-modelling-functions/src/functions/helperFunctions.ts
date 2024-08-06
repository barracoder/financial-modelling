
export const inputWorksheetName = "Input";

// Function to convert Excel date to JavaScript Date
export function excelDateToJSDate(excelDate: number): Date {
  const date = new Date((excelDate - (25567 + 1)) * 86400 * 1000);
  return date;
}

// Function to check if a date is in a specific year
export async function isDateInYear(excelDate: number, year: number): Promise<boolean> {
  const date = await excelDateToJSDate(excelDate);
  const test = date.getFullYear() === year;
  return test;
}

export function isModelWorksheet(context: Excel.RequestContext) {
  const isModel = context.workbook.worksheets.getItemOrNullObject("Input");
  const modelExists = isModel.load("isNullObject");
  context.sync();

  return modelExists;
}
