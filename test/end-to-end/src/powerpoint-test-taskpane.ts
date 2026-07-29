import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { runPowerPoint } from "../../../src/taskpane/powerpoint";
import * as testHelpers from "./test-helpers";

/* global Office, Promise */

const port: number = 4201;
let testValues: any = [];

Office.onReady(async (info) => {
  if (info.host === Office.HostType.PowerPoint) {
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

async function getText(): Promise<string> {
  return PowerPoint.run(async (context: PowerPoint.RequestContext) => {
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
    await runPowerPoint();
    await testHelpers.sleep(2000);

    const text = await getText();

    testHelpers.addTestResult(testValues, "output-message", text, "Hello World!");
    await sendTestResults(testValues, port);
  } catch (err) {
    testValues = [];
    testHelpers.addErrorResult(testValues, `runTest failed: ${testHelpers.formatError(err)}`);
    await sendTestResults(testValues, port).catch(() => {});
  }
}
