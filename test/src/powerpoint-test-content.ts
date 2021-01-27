import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { run } from "../../src/content/powerpoint";
import * as testHelpers from "./test-helpers";
const port: number = 4201;
let testValues: any = [];

Office.onReady(async (info) => {
    if (info.host === Office.HostType.PowerPoint) {
        const testServerResponse: object = await pingTestServer(port);
        if (testServerResponse["status"] == 200) {
            await runTest();
        }
    }
});

export async function runTest(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        try {
            // Execute taskpane code
            run().then(() => {
                run().then(() => {
                    testHelpers.sleep(6000).then(() => {
                        // Get output of executed taskpane code
                        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async (asyncResult) => {
                            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                                console.error(asyncResult.error.message);
                                testHelpers.addTestResult(testValues, "output-message", asyncResult.error.message, " Hello World!");
                            } else {
                                console.log(`The selected data is "${asyncResult.value}".`);
                                testHelpers.addTestResult(testValues, "output-message", asyncResult.value, " Hello World!");
                            }
                            await sendTestResults(testValues, port);
                            testValues.pop();
                            resolve();
                        });
                    });
                });
            });
        } catch {
            reject();
        }
    });
}