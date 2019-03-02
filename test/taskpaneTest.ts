import * as testHelper from "../test/testHelpers";
const port: number = 8080;
let testValues = []; 

export async function isTestServerStarted(): Promise<boolean> {
    const testServerResponse: any = await testHelper.pingTestServer(port);
    if (testServerResponse["status"] === 200) {
        return true;
    }
    else {
        return false;
    }
}

export async function readSendData(): Promise<void> {
    await Excel.run(async context => {
            const range = context.workbook.getSelectedRange();
            const cellFill = range.format.fill;
            cellFill.load('color');
            await context.sync();

            var data  = {};
            var nameKey = "Name";
            var valueKey = "Value";            
            data[nameKey] = "fill-color";
            data[valueKey] = cellFill.color;
            testValues.push(data);

            if (testValues.length > 0) {
                testHelper.sendTestResults(testValues, port)
            }
    });
}
