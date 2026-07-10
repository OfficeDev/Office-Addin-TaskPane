import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { runPowerPoint } from "../../../src/taskpane/powerpoint";
import * as testHelpers from "./test-helpers";

/* global Office, Promise */

const port: number = 4201;
let testValues: any = [];
const steps: string[] = [];

Office.onReady(async (info) => {
  steps.push(`Office.onReady: host=${info.host}, platform=${info.platform}`);
  if (info.host === Office.HostType.PowerPoint) {
    try {
      steps.push("pingTestServer");
      const testServerResponse: object = await pingTestServer(port);
      if (testServerResponse["status"] == 200) {
        steps.push("ping OK, running test");
        await runTest();
      } else {
        steps.push(`ping returned unexpected status: ${JSON.stringify(testServerResponse)}`);
        testHelpers.addTestResult(testValues, "test-error", `Ping failed: ${JSON.stringify(testServerResponse)}`, "no-error");
        testHelpers.addDiagnosticResult(testValues, steps);
        await sendTestResults(testValues, port).catch(() => {});
      }
    } catch (err) {
      steps.push(`initialization error: ${err}`);
      testHelpers.addTestResult(testValues, "test-error", `Initialization failed: ${err}`, "no-error");
      testHelpers.addDiagnosticResult(testValues, steps);
      await sendTestResults(testValues, port).catch(() => {});
    }
  }
});

async function getText(): Promise<string> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load("items");
    await context.sync();

    const textboxShapes = shapes.items.filter(shape => shape.name === "GreetingTextbox");
      
    if (textboxShapes.length > 0) {
      const textFrame = textboxShapes[0].textFrame.load("textRange");
      await context.sync();

      return textFrame.textRange.text;
    } else {
      return "";
    }
  });
}

export async function runTest(): Promise<void> {
  try {
    steps.push("runPowerPoint start");
    await runPowerPoint();
    steps.push("runPowerPoint complete");
    await testHelpers.sleep(2000);

    steps.push("getText start");
    const text = await getText();
    steps.push(`getText result: "${text}"`);

    testHelpers.addTestResult(testValues, "output-message", text, "Hello World!");
    testHelpers.addDiagnosticResult(testValues, steps);
    await sendTestResults(testValues, port);
  } catch (err) {
    steps.push(`runTest error: ${err}`);
    testValues = [];
    testHelpers.addTestResult(testValues, "test-error", `runTest failed: ${err}`, "no-error");
    testHelpers.addDiagnosticResult(testValues, steps);
    await sendTestResults(testValues, port).catch(() => {});
  }
}
