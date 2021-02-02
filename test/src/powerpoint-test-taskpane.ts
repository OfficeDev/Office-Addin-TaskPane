import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { run } from "../../src/taskpane/powerpoint";
import * as testHelpers from "./test-helpers";
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

export async function runTest(): Promise<void> {
    // Set up textbox and cursor for taskpane code test
    await new Promise<void>((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(
            " ",
            {
                coercionType: Office.CoercionType.Text
            },
            result => {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error(result.error.message);
                    reject(result.error);
                }
                resolve();
            }
        )
    });

    // Execute taskpane code
    await run();
    await testHelpers.sleep(6000);
    await actualTest();
}

async function actualTest() {
    // Get output of executed taskpane code
    return new Promise<void>((resolve) => {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
                testHelpers.addTestResult(testValues, "output-message", asyncResult.error.message, "Hello World!");
            } else {
                console.log(`The selected data is "${asyncResult.value}".`);
                testHelpers.addTestResult(testValues, "output-message", asyncResult.value, "Hello World!");
            }
            await sendTestResults(testValues, port);
            testValues.pop();
            resolve();
        });
    });
}