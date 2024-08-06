import * as Helpers from "./helperFunctions";

let cachedInputRange: any[][] = [];
const queue: Array<{ key: any; year: any; resolve: (value: number) => void; reject: (reason?: any) => void }> = [];
let isProcessing = false;
let cachedYears: any[] = [];
let cachedKeys: any[] = [];
const keyOffset = 1;
const yearOffset = 2;
const yearRow = 0;
const inputDataRangeAddress = "Input!A7:AA500";
let keyYearDictionary: Map<string, number> = new Map();

function buildKey(key: string, year: number): string {
  return `${key}/${year}`;
}

function getInputWorksheet(context: Excel.RequestContext): Excel.Worksheet {
  return context.workbook.worksheets.getItem(Helpers.inputWorksheetName);
}

async function createKeyYearDictionary(values: any[][]): Promise<Map<string, number>> {
  try {
    const years = values[0].slice(yearOffset).map((date) => Helpers.excelDateToJSDate(date).getFullYear()); // Extract years from the top row, excluding the first cell
    const xyDictionary = new Map<string, number>();

    for (let keyIndex = 1; keyIndex < values.length; keyIndex++) {
      const key = values[keyIndex][keyOffset]; // Extract key from the first column
      if (!key) continue;
      for (let yearIndex = 0; yearIndex < years.length; yearIndex++) {
        const year = years[yearIndex];
        const value = values[keyIndex][yearIndex + yearOffset];
        xyDictionary.set(buildKey(key, year), value);
      }
    }

    console.log(xyDictionary);
    return xyDictionary;
  } catch (error) {
    console.error(error);
    throw error;
  }
}

/**
 * Caches the values from the Input Excel worksheet.
 *
 */
function _populateCaches() {
  console.log("Populating caches");
  Excel.run(async (context) => {
    if(!Helpers.isModelWorksheet(context)) return;
    const inputDataRange = context.workbook.worksheets.getItem("Input").getRange(inputDataRangeAddress);
    inputDataRange.load("values");

    await context.sync();

    const values = inputDataRange.values;
    cachedInputRange = values;
    cachedYears = values[yearRow].slice(yearOffset); // First row for years
    cachedKeys = values.map((row) => row[keyOffset]); // First column for keys
    keyYearDictionary = await createKeyYearDictionary(values);
    console.log(cachedInputRange, cachedKeys, cachedYears, keyYearDictionary);
  });
}

Office.onReady(() => {
  _populateCaches();
});

export function enqueueLookup(key: string, year: number, resolve: (value: number) => void, reject: (reason?: any) => void) {
  queue.push({ key, year, resolve, reject });
  processQueue();
}

async function processQueue() {
  if (isProcessing || queue.length === 0) {
    return;
  }

  isProcessing = true;
  const { key, year, resolve, reject } = queue.shift()!;

  try {
    if (!cachedInputRange) {
      await _populateCaches();
    }
    
    // Return the value found at the intersection of the key row and year column
    const mapKey = buildKey(key, year);

    if (!keyYearDictionary.has(mapKey)) {
      const message = `The key "${key}" and year "${year}" were not found in the table.`;
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, message);
    }
    const result = keyYearDictionary.get(mapKey);

    resolve(result);
  } catch (error) {
    console.log(error);
    reject(error);
  } finally {
    isProcessing = false;
    processQueue();
  }
}

async function onInputWorksheetChanged(args: Excel.WorksheetChangedEventArgs) {
  console.log("Input worksheet changed", args.address, args.details, args.type, args.changeDirectionState);
}

Office.onReady(() => {
  Excel.run(async (context) => {
    if(!Helpers.isModelWorksheet(context)) return;
    const inputWorksheet = context.workbook.worksheets.getItem("Input");
    inputWorksheet.onChanged.add(onInputWorksheetChanged);

    await context.sync();
    console.log("Event handler successfully registered for onChanged event in the Input worksheet.");
  }).catch((error) => {
    console.error(error);
  });
});

