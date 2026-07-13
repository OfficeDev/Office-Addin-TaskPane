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

async function getDocumentState(): Promise<string> {
  const info: string[] = [];
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.load("text,type");
      await context.sync();
      info.push(`bodyText="${body.text.substring(0, 100)}"`);
      info.push(`bodyType="${body.type}"`);
    });
  } catch (err) {
    info.push(`body read failed: ${testHelpers.formatError(err)}`);
  }

  // Try desktop-only properties separately to avoid crashing if unsupported
  try {
    await Word.run(async (context) => {
      const doc = context.document;
      doc.load("saved");
      await context.sync();
      info.push(`saved=${doc.saved}`);
    });
  } catch {
    info.push("saved=unsupported");
  }

  try {
    await Word.run(async (context) => {
      const doc = context.document;
      doc.load("isReadOnly");
      await context.sync();
      info.push(`isReadOnly=${doc.isReadOnly}`);
    });
  } catch {
    info.push("isReadOnly=unsupported");
  }

  try {
    await Word.run(async (context) => {
      const doc = context.document;
      doc.load("protectionType");
      await context.sync();
      info.push(`protectionType=${doc.protectionType}`);
    });
  } catch {
    info.push("protectionType=unsupported");
  }

  try {
    await Word.run(async (context) => {
      const doc = context.document;
      doc.load("fullName");
      await context.sync();
      info.push(`fullName="${doc.fullName}"`);
    });
  } catch {
    info.push("fullName=unsupported");
  }

  return info.join(", ");
}

export async function runTest() {
  try {
    // Gather document state info for diagnostics
    const docState = await getDocumentState();
    steps.push(`document state: ${docState}`);

    steps.push("runWord start");
    await runWord();
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
