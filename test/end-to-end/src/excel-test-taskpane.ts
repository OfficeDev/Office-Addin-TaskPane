import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { runExcel } from "../../../src/taskpane/excel";
import * as testHelpers from "./test-helpers";

/* global Excel, Office */

const port: number = 4201;
let testValues: any = [];
const steps: string[] = [];

Office.onReady(async (info) => {
  steps.push(`Office.onReady: host=${info.host}, platform=${info.platform}`);
  if (info.host === Office.HostType.Excel) {
    try {
      steps.push("pingTestServer");
      const testServerResponse: object = await pingTestServer(port);
      if (testServerResponse["status"] == 200) {
        steps.push("ping OK, running test");
        await runTest();
      } else {
        steps.push(`ping returned unexpected status: ${JSON.stringify(testServerResponse)}`);
        testHelpers.addErrorResult(testValues, `Ping failed: ${JSON.stringify(testServerResponse)}`);
        testHelpers.addDiagnosticResult(testValues, steps);
        await sendTestResults(testValues, port).catch(() => {});
      }
    } catch (err) {
      steps.push(`initialization error: ${testHelpers.formatError(err)}`);
      testHelpers.addErrorResult(testValues, `Initialization failed: ${testHelpers.formatError(err)}`);
      testHelpers.addDiagnosticResult(testValues, steps);
      await sendTestResults(testValues, port).catch(() => {});
    }
  }
});

export async function runTest(): Promise<void> {
  try {
    steps.push("runExcel start");
    await runExcel();
    steps.push("runExcel complete");
    await testHelpers.sleep(2000);

    steps.push("Excel.run start");
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const cellFill = range.format.fill;
      cellFill.load("color");
      await context.sync();
      steps.push("context.sync complete");
      await testHelpers.sleep(2000);

      steps.push(`fill color: ${cellFill.color}`);
      testHelpers.addTestResult(testValues, "fill-color", cellFill.color, "#FFFF00");
      testHelpers.addDiagnosticResult(testValues, steps);
      await sendTestResults(testValues, port);
      testValues.pop();
      testValues.pop();
      await testHelpers.closeWorkbook();
    });
  } catch (err) {
    steps.push(`runTest error: ${testHelpers.formatError(err)}`);
    testValues = [];
    testHelpers.addErrorResult(testValues, `runTest failed: ${testHelpers.formatError(err)}`);
    testHelpers.addDiagnosticResult(testValues, steps);
    await sendTestResults(testValues, port).catch(() => {});
  }
}
