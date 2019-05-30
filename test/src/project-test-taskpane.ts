import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { run } from "../../src/taskpane/project";
import * as testHelpers from "./test-helpers";
const port: number = 4201;
let testValues: any = [];

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Project) {
        const testServerResponse: object = await pingTestServer(port);
        if (testServerResponse["status"] == 200) {
            await runTest();
        }
    }
});


export async function runTest(): Promise<void> {
    return new Promise<void>(async (resolve, reject) => {
        /**
        * Insert your Project test code here
        */
    });
}