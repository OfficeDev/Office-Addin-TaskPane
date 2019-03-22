import * as testHelper from "office-addin-test-helpers"; 
const port: number = 8080;

export async function isTestServerStarted(): Promise<boolean> {
    return new Promise<boolean>(async (resolve, reject) => {
        const testServerResponse: any = await testHelper.pingTestServer(port);
        if (testServerResponse["status"] === 200) {
            resolve(true);
        } else {
            resolve(false);
        }
    });
}

export async function runTest(host: string) {
    switch (host.toLocaleLowerCase()) {
        case 'excel':
            runExcelTest();
            break;
        case 'onenote':
            runOneNoteTest();
            break;
        case 'outlook':
            runOutlookTest();
            break;
        case 'powerpoint':
            runPowerPointTest();
            break;
        case 'project':
            runProjectTest();
            break;
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
        sendTestResults(cellFill.color, "fill-color");
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

async function runPowerPointTest(): Promise<void>{
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
        sendTestResults(result.value, "test-string")
    }); 
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

async function sendTestResults(result: any, resultType: string): Promise<void> {
    var data = {};
    let testValues = [];
    var nameKey = "Name";
    var valueKey = "Value";
    data[nameKey] = resultType;
    data[valueKey] = result
    testValues.push(data);

    await testHelper.sendTestResults(testValues, port);
}
