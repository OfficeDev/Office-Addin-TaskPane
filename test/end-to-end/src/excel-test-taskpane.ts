import { pingTestServer, sendTestResults, TestServerResponse } from "office-addin-test-helpers";
import { runExcel } from "../../../src/shared/excel";
import * as testHelpers from "./test-helpers";

/* global Excel Office */

const port: number = 4201;
let testValues: any = [];

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    const testServerResponse: TestServerResponse = await pingTestServer(port);
    if (testServerResponse["status"] == 200) {
      await runTest();
    }
  }
});

export async function runTest(): Promise<void> {
  // Execute taskpane code
  await runExcel();
  await testHelpers.sleep(2000);

  // Get output of executed taskpane code
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const cellFill = range.format.fill;
    cellFill.load("color");
    await context.sync();
    await testHelpers.sleep(2000);

    testHelpers.addTestResult(testValues, "fill-color", cellFill.color, "#FFFF00");
    await sendTestResults(testValues, port);
    testValues.pop();
    await testHelpers.closeWorkbook();
  });
}
