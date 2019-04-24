import * as excelTaskpane from '../taskpane/excel';
import * as testDataJson from './testData.json';
import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
const port: number = 4201;
const testData = (<any>testDataJson).hosts;
let testValues = [];

Office.initialize = async () => {
    document.getElementById('sideload-msg').style.display = 'none';
    document.getElementById('app-body').style.display = 'flex';
    document.getElementById('run').onclick = run;

    const testServerResponse: object = await pingTestServer(port);
    if (testServerResponse["status"] == 200) {
        await runTest(testServerResponse["platform"]);
        await sendTestResults(testValues, port);
    }
};

export async function runTest(platform: string) {
    switch (Office.context.host.toString().toLowerCase()) {
        case 'excel':
            await runExcelTest(platform);
            break;
        case 'onenote':
            await runOneNoteTest(platform);
            break;
        case 'outlook':
            await runOutlookTest(platform);
            break;
        case 'powerpoint':
            await runPowerPointTest(platform);
            break;
        case 'project':
            await runProjectTest(platform);
            break;
        case 'word':
            await runWordTest(platform);
    }
}

async function runExcelTest(platform: string): Promise<void> {
    return new Promise<void>(async (resolve, reject) => {
        try {
            // Execute taskpane code
            await excelTaskpane.run();

            // Mac is much slower so we need to wait longer for the function to return a value
            await sleep(platform === "Win32" ? 2000 : 8000);

            // Get output of executed taskpane code
            await Excel.run(async context => {
                const range = context.workbook.getSelectedRange();
                const cellFill = range.format.fill;
                cellFill.load('color');
                await context.sync();

                // Mac is much slower so we need to wait longer for the function to return a value
                await sleep(platform === "Win32" ? 2000 : 8000);

                addTestResult(testData.excel.resultName, cellFill.color);
                resolve();
            });
        } catch {
            reject();
        }
    });
}

async function runOneNoteTest(platform: string) {
    /**
     * Insert your Outlook code here
     */
}

async function runOutlookTest(platform: string) {
    /**
     * Insert your Outlook code here
     */
}

async function runPowerPointTest(platform: string): Promise<void>{
    /**
     * Insert your Outlook code here
     */
}

async function runProjectTest(platform: string) {
    /**
     * Insert your Outlook code here
     */
}

async function runWordTest(platform: string) {
    /**
     * Insert your Outlook code here
     */
}

function addTestResult(resultName: string, resultValue: any) {
    var data = {};
    var nameKey = "Name";
    var valueKey = "Value";
    data[nameKey] = resultName;
    data[valueKey] = resultValue;
    testValues.push(data);
}

async function sleep(ms: number): Promise<any> {
    return new Promise(resolve => setTimeout(resolve, ms));
}
