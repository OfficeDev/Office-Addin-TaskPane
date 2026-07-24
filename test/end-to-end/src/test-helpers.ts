import * as childProcess from "child_process";

/* global Excel, process, Promise, setTimeout */

export async function closeDesktopApplication(application: string): Promise<boolean> {
  let processName: string = "";
  switch (application.toLowerCase()) {
    case "excel":
      processName = "Excel";
      break;
    case "powerpoint":
      processName = process.platform === "win32" ? "Powerpnt" : "PowerPoint";
      break;
    case "onenote":
      processName = "Onenote";
      break;
    case "outlook":
      processName = "Outlook";
      break;
    case "project":
      processName = "Project";
      break;
    case "word":
      processName = process.platform === "win32" ? "Winword" : "Word";
      break;
    default:
      throw new Error(`${application} is not a valid Office desktop application.`);
  }

  await sleep(3000); // wait for host to settle
  try {
    let cmdLine: string;
    if (process.platform == "win32") {
      cmdLine = `tskill ${processName}`;
    } else {
      cmdLine = `pkill ${processName}`;
    }

    return await executeCommandLine(cmdLine);
  } catch (err) {
    throw new Error(`Unable to kill ${application} process. ${err}`);
  }
}

export async function closeWorkbook(): Promise<void> {
  await sleep(1000);
  await Excel.run(async (context: Excel.RequestContext) => context.workbook.close(Excel.CloseBehavior.skipSave));
}

export function addTestResult(testValues: any[], resultName: string, resultValue: any, expectedValue: any) {
  testValues.push({
    expectedValue,
    resultName,
    resultValue,
  });
}

export function addErrorResult(testValues: any[], errorMessage: string) {
  testValues.push({
    resultName: "test-error",
    resultValue: errorMessage,
    expectedValue: "no-error",
  });
}

export function formatError(err: any): string {
  if (!err) return "Unknown error (null/undefined)";

  const parts: string[] = [];

  // Basic message
  if (err.message) {
    parts.push(`Message: ${err.message}`);
  } else {
    parts.push(`${err}`);
  }

  // Office.js error code
  if (err.code) {
    parts.push(`Code: ${err.code}`);
  }

  // Office.js debugInfo (OfficeExtension.Error)
  if (err.debugInfo) {
    if (err.debugInfo.code) parts.push(`DebugCode: ${err.debugInfo.code}`);
    if (err.debugInfo.message) parts.push(`DebugMessage: ${err.debugInfo.message}`);
    if (err.debugInfo.errorLocation) parts.push(`Location: ${err.debugInfo.errorLocation}`);
    if (err.debugInfo.innerError) {
      const inner = err.debugInfo.innerError;
      parts.push(`InnerError: ${inner.code || ""} ${inner.message || JSON.stringify(inner)}`);
    }
  }

  // Stack trace
  if (err.stack) {
    parts.push(`Stack: ${err.stack}`);
  }

  return parts.join(" | ");
}

export async function sleep(ms: number): Promise<any> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function executeCommandLine(cmdLine: string): Promise<boolean> {
  return new Promise<boolean>((resolve, reject) => {
    childProcess.exec(cmdLine, (error) => {
      if (error) {
        reject(false);
      } else {
        resolve(true);
      }
    });
  });
}
