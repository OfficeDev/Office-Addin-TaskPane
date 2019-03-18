import * as testHelper from "office-addin-test-helpers";
const port: number = 8080;
let testValues = [];

export async function isTestServerStarted(): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        const testServerResponse: any = await testHelper.pingTestServer(port);
        if (testServerResponse["status"] === 200) {
            resolve(true);
        }
        else {
            resolve(false);
        }
    });
}

export async function runTest(host: string) {
    switch (host.toLocaleLowerCase()) {
        case 'excel':
            runExcelTest();
        case 'onenote':
            runOneNoteTest();
        case 'outlook':
            runOutlookTest();
        case 'powerpoint':
            runPowerPointTest()
        case 'project':
            runProjectTest();
        case 'word':
            return runWordTest();
    }
}

async function runExcelTest(): Promise<void> {
    await Excel.run(async context => {
        const range = context.workbook.getSelectedRange();
        const cellFill = range.format.fill;
        cellFill.load('color');
        await context.sync();

        var data = {};
        var nameKey = "Name";
        var valueKey = "Value";
        data[nameKey] = "fill-color";
        data[valueKey] = cellFill.color;
        testValues.push(data);

        if (testValues.length > 0) {
            sendTestResults();
        }
    });
}

async function runOneNoteTest() {
    /**
     * Insert your Outlook code here
     */
}

async function runOutlookTest() {
    /**
     * Insert your Outlook code here
     */
}

async function runPowerPointTest() {
    /**
     * Insert your Outlook code here
     */
}

async function runProjectTest() {
    /**
     * Insert your Outlook code here
     */
}

async function runWordTest() {
    /**
     * Insert your Outlook code here
     */
}

async function sendTestResults(): Promise<void> {
    await testHelper.sendTestResults(testValues, port);
}