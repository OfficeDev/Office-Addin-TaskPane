import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { run } from "../../../src/taskpane/powerpoint";
import * as testHelpers from "./test-helpers";

/* global Office, Promise */

const port: number = 4201;
let testValues: any = [];

Office.onReady(async (info) => {
  if (info.host === Office.HostType.PowerPoint) {
    const testServerResponse: object = await pingTestServer(port);
    if (testServerResponse["status"] == 200) {
      runTest();
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
  // Execute taskpane code
  await run();
  await testHelpers.sleep(2000);

  // get inserted selected text
  const text = await getText();

  // send test results
  testHelpers.addTestResult(testValues, "output-message", text, "Hello World!");

  await sendTestResults(testValues, port);
}
