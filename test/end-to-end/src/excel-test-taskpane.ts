import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { runExcel } from "../../../src/taskpane/excel";
import * as testHelpers from "./test-helpers";

/* global Excel, Office */

const port: number = 4201;
let testValues: any = [];

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    try {
      const testServerResponse: object = await pingTestServer(port);
      if (testServerResponse["status"] == 200) {
        await runTest();
      } else {
        testHelpers.addErrorResult(testValues, `Ping failed: ${JSON.stringify(testServerResponse)}`);
        await sendTestResults(testValues, port).catch(() => {});
      }
    } catch (err) {
      testHelpers.addErrorResult(testValues, `Initialization failed: ${testHelpers.formatError(err)}`);
      await sendTestResults(testValues, port).catch(() => {});
    }
  }
});

export async function runTest(): Promise<void> {
  try {
    await runExcel();
    await testHelpers.sleep(2000);

    await Excel.run(async (context) => {
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
  } catch (err) {
    testValues = [];
    testHelpers.addErrorResult(testValues, `runTest failed: ${testHelpers.formatError(err)}`);
    await sendTestResults(testValues, port).catch(() => {});
  }
}
