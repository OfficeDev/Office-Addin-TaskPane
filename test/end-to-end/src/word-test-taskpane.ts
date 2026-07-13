import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { runWord } from "../../../src/taskpane/word";
import * as testHelpers from "./test-helpers";

/* global Office, Word */

const port: number = 4201;
let testValues: any = [];
const steps: string[] = [];

Office.onReady(async (info) => {
  steps.push(`Office.onReady: host=${info.host}, platform=${info.platform}`);
  if (info.host === Office.HostType.Word) {
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

async function retryOperation(operation: () => Promise<void>, maxRetries: number = 5, delayMs: number = 3000): Promise<void> {
  for (let i = 0; i < maxRetries; i++) {
    try {
      await operation();
      return;
    } catch (err) {
      steps.push(`retry ${i + 1}/${maxRetries} failed: ${testHelpers.formatError(err)}`);
      if (i === maxRetries - 1) throw err;
      await testHelpers.sleep(delayMs);
    }
  }
}

export async function runTest() {
  try {
    steps.push("runWord start");
    await retryOperation(() => runWord());
    steps.push("runWord complete");
    await testHelpers.sleep(2000);

    steps.push("Word.run start");
    return Word.run(async (context) => {
      var firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.load("text");
      await context.sync();
      steps.push("context.sync complete");
      await testHelpers.sleep(2000);

      steps.push(`paragraph text: "${firstParagraph.text}"`);
      testHelpers.addTestResult(testValues, "output-message", firstParagraph.text, "Hello World");
      testHelpers.addDiagnosticResult(testValues, steps);
      await sendTestResults(testValues, port);
      testValues.pop();
      testValues.pop();
    });
  } catch (err) {
    steps.push(`runTest error: ${testHelpers.formatError(err)}`);
    testValues = [];
    testHelpers.addErrorResult(testValues, `runTest failed: ${testHelpers.formatError(err)}`);
    testHelpers.addDiagnosticResult(testValues, steps);
    await sendTestResults(testValues, port).catch(() => {});
  }
}
