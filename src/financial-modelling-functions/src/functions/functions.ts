/* eslint-disable @typescript-eslint/no-unused-vars */
/* global clearInterval, console, CustomFunctions, setInterval */

// import { enqueueLookup } from "./getFinancials";
import * as Helpers from "./helperFunctions";

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function addTwoNumbers3(first: number, second: number): number {
  return first + second;
}
// /**
//  * Looks up a value in the Input Excel worksheet.
//  * @customfunction
//  * @param key Key to lookup in the table.
//  * @param year Year to lookup in the table.
//  * @returns The value found in the table.
//  * @volatile
//  */
// async function getFinancials(key: string, year: number): Promise<number | unknown> {
//   console.log("getFinancials", key, year);
//   return new Promise((resolve, reject) => {
//     enqueueLookup(key, year, resolve, reject);
//   }).then((result) => {
//     return result;
//   }).catch((error) => {
//     console.error(error);
//   });
// }


/**
 * Retrieves the value from a table based on the specified row and column keys.
 * @customfunction
 * @param {string} rowKey The key to look for in the first column.
 * @param {number} columnYear The year to look for in the header row.
 * @param {string} tableRange The address of the table range.
 * @returns The value in the table that matches the row and column keys.
 */
function getValueFromTable(rowKey, columnYear) {
  Excel.run(async (context) => {
    
    const sheet = context.workbook.worksheets.getItem("Input")
    const table = sheet.tables.getItem("FinancialsData");

    // Load the table's headers and data.
    table.columns.load("items/name");
    table.rows.load("items/values");
    await context.sync();

    console.log("getValueFromTable", rowKey, columnYear);
    
    console.log("getValueFromTable", rowKey, columnYear);
    // Find the column index for the given year.
    const columnHeaders = table.columns.items.map(col => col.name);
    console.log(columnHeaders);
    const columnIndex = columnHeaders.indexOf(columnYear);

    if (columnIndex === -1) {
      const message = `Column year ${columnYear} not found. Please check that the year is in the correct format.`;
      console.log(message);
      return new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }

    // Find the row index for the given key.
    const rows = table.rows.items;
    const rowIndex = rows.findIndex(row => row.values[0][0] === rowKey);

    if (rowIndex === -1) {
      const message = `Row key ${rowKey} not found. Please check that the key is in the correct format.`;
      console.log(message);
      return new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }

    // Return the cell value.
    const cellValue = rows[rowIndex].values[0][columnIndex];
    return cellValue;
  });
}
