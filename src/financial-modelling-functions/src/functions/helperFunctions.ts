
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

export function formatExcelDate(excelDate, format) {
  // Excel dates are the number of days since January 1, 1900.
  const excelEpoch = new Date(1899, 11, 30);
  const date = new Date(excelEpoch.getTime() + excelDate * 86400000);

  // Define the options for date formatting.
  const options = {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  };

  // Format the date based on the given format.
  switch (format) {
    case 'dd/MM/yyyy':
      return date.toLocaleDateString('en-GB', options);
    case 'MM/dd/yyyy':
      return date.toLocaleDateString('en-US', options);
    case 'yyyy-MM-dd':
      return date.toISOString().split('T')[0];
    default:
      return date.toLocaleDateString('en-GB', options);
  }
}
