/* global clearInterval, console, CustomFunctions, setInterval */

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
 * Adds two numbers and wait 10s return.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function addAndWait10s(first: number, second: number): number {
  wait(10000);
  return first + second;
}

/**
 * Adds two numbers and wait 2s return.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function addAndWait2s(first: number, second: number): number {
  wait(2000);
  return first + second;
}

/**
 * Adds two numbers and wait 3s return.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function addAndWait3s(first: number, second: number): number {
  wait(3000);
  return first + second;
}

/**
 * Adds two numbers and wait 5s return.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function addAndWait5s(first: number, second: number): number {
  wait(5000);
  return first + second;
}

function wait(ms) {
  var start = Date.now(), now = start;
  while (now - start < ms) {
    now = Date.now();
  }
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
 * Calculates the volume of a sphere.
 * @customfunction
 * @returns The volume of the sphere.
 */
async function jiaruitestcf2() {
  // Retrieve the context object.
  const context = new Excel.RequestContext();

  // Use the context object to access the cell at the input address.
  const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  range.load("values");
  await context.sync();

  // Return the value of the cell at the input address.
  return range.values[0][0];
}


/**
 * Take a number as the input value and return a formatted number value as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function createFormattedNumber(value, format) {
  return {
      type: "FormattedNumber",
      basicValue: value,
      numberFormat: format
  }
}