/* eslint-disable @typescript-eslint/no-unused-vars */
/* global clearInterval, console, CustomFunctions, setInterval */

// import { enqueueLookup } from "./getFinancials";
import * as Helpers from "./helperFunctions";

const queue: Array<{ key: any; year: any; resolve: (value: number) => void; reject: (reason?: any) => void }> = [];
let isProcessing = false;

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

/**
 * @customfunction
 * @param key The string key to lookup in the Financials table
 * @param year The year to lookup in the Financials table
 * @returns
 * @volatile
 */
export async function GetInputData(key, year) {
  return new Promise((resolve, reject) => {
    queue.push({ key, year, resolve, reject });
    processQueue();
  });
}

/**
 * The function that does the actual processing of the queue of table lookup calls
 * @returns 
 */
async function processQueue() {
  if (isProcessing || queue.length === 0) {
    return;
  }

  isProcessing = true;

  try {
    await Excel.run(async (context) => {
      console.log("Getting data from context");
      const table = context.workbook.tables.getItem("FinancialsData");
      const headers = table.getHeaderRowRange().load("values");
      const keysRange = table.getDataBodyRange().getColumn(0).load("values");
      const dataBodyRange = table.getDataBodyRange().load("values");

      await context.sync();

      while (queue.length > 0) {
        console.log("Processing queue. Remaining items: ", queue.length);
        const { key, year, resolve, reject } = queue.shift()!;

        let errorMessage;

        if (key === undefined || year === undefined || year === null || key === null) {
          errorMessage = `Key ${key} or year ${year} is undefined`;
          console.log(errorMessage);
          reject(new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, errorMessage));
          continue;
        }

        const headerValues = headers.values[0];
        const keysValues = keysRange.values.map((row) => row[0]);

        const yearIndex = headerValues.indexOf(year.toString());
        const keyIndex = keysValues.indexOf(key);

        // console.log("Header Values", headerValues);
        // console.log("Keys Values", keysValues);
        // console.log("Year Index", yearIndex);
        // console.log("Key Index", keyIndex);

        if (yearIndex === -1) {
          reject(new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, `Year ${year} not found in headers`));
          continue;
        }

        if (keyIndex === -1) {
          reject(new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, `Key ${key} not found in keys`));
          continue;
        }

        resolve(dataBodyRange.values[keyIndex][yearIndex]);
      }
    });
  } catch (error) {
    console.error("Error processing queue", error);
  } finally {
    isProcessing = false;
    if (queue.length > 0) {
      processQueue(); // Process remaining items in the queue
    }
  }
}
