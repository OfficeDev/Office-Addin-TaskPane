import * as excelTaskpane from './../../src/taskpane/excel'
import * as wordTaskpane from './../../src/taskpane/word';
import * as testDataJson from './test-data.json';
import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
const port: number = 4201;
const testData = (<any>testDataJson).hosts;
let host = undefined;
let testValues = [];

Office.initialize = async () => {
    document.getElementById('sideload-msg').style.display = 'none';
    document.getElementById('app-body').style.display = 'flex';
    document.getElementById('run').onclick = run;

    const testServerResponse: object = await pingTestServer(port);
    if (testServerResponse["status"] == 200) {
        await runTest(testServerResponse["platform"]);
        await sendTestResults(testValues, port);
        testValues.pop();
        if (host === "excel") {
            await closeWorkbook();
        }
    }
};

export async function runTest(platform: string) {
    switch (Office.context.host.toString().toLowerCase()) {
        case 'excel':
            host = "excel";
            await runExcelTest(platform);
            break;
        case 'onenote':
            host = "onenote";
            await runOneNoteTest(platform);
            break;
        case 'outlook':
            host = "outlook";
            await runOutlookTest(platform);
            break;
        case 'powerpoint':
            host = "powerpoint";
            await runPowerPointTest(platform);
            break;
        case 'project':
            host = "project";
            await runProjectTest(platform);
            break;
        case 'word':
            host = "word";
            await runWordTest(platform);
    }
}

async function runExcelTest(platform: string): Promise<void> {
    return new Promise<void>(async (resolve, reject) => {
        try {
            // Execute taskpane code
            await excelTaskpane.run();
            await sleep(2000);

            // Get output of executed taskpane code
            await Excel.run(async context => {
                const range = context.workbook.getSelectedRange();
                const cellFill = range.format.fill;
                cellFill.load('color');
                await context.sync();
                await sleep(2000);

                addTestResult(testData.Excel.resultName, cellFill.color);
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

async function runPowerPointTest(platform: string): Promise<void> {
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
    return new Promise<void>(async (resolve, reject) => {
        try {
            // Execute taskpane code
            await wordTaskpane.run();
            await sleep(2000);

            // Get output of executed taskpane code
            Word.run(async (context) => {
                var firstParagraph = context.document.body.paragraphs.getFirst();
                firstParagraph.load("text");
                await context.sync();
                await sleep(2000);

                addTestResult(testData.Word.resultName, firstParagraph.text);
                resolve();
            });
        } catch {
            reject();
        }
    });
}

async function closeWorkbook(): Promise<void> {
    return new Promise<void>(async (resolve, reject) => {
        try {
            await Excel.run(async context => {
                // @ts-ignore
                context.workbook.close(Excel.CloseBehavior.skipSave);
                resolve();
            });
        } catch {
            reject();
        }
    });
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
