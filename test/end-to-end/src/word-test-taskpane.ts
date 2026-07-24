import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { runWord } from "../../../src/taskpane/word";
import * as testHelpers from "./test-helpers";

/* global Office, Word */

const port: number = 4201;
let testValues: any = [];

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    try {
      const testServerResponse = (await pingTestServer(port)) as { status?: number };
      if (testServerResponse.status == 200) {
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

export async function runTest() {
  try {
    await runWord();
    await testHelpers.sleep(2000);

    await Word.run(async (context: Word.RequestContext) => {
      var firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.load("text");
      await context.sync();
      await testHelpers.sleep(2000);

      testHelpers.addTestResult(testValues, "output-message", firstParagraph.text, "Hello World");
      await sendTestResults(testValues, port);
      testValues.pop();
    });
  } catch (err) {
    testValues = [];
    testHelpers.addErrorResult(testValues, `runTest failed: ${testHelpers.formatError(err)}`);
    await sendTestResults(testValues, port).catch(() => {});
  }
}
