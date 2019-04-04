import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from "path";
const manifestPath = "./src/test/test-manifest.xml";
const testJsonFile: string = path.resolve(`${process.cwd()}/src/test/testData.json`);
const testJsonData = JSON.parse(fs.readFileSync(testJsonFile).toString());
import * as testHelper from "office-addin-test-helpers"
import * as testServerInfra from "office-addin-test-server";
const port: number = 4201;
const testServer = new testServerInfra.TestServer(port);

// Only run tests on Windows for now until the Close Workbook API is enabled in Production
if (process.platform == 'win32') {
    Object.keys(testJsonData.hosts).forEach(function (host) {
        const resultName = testJsonData.hosts[host].resultName;
        const resultValue: string = testJsonData.hosts[host].resultValue;
        let testValues: any = [];

        describe("Test Task Pane Project", function () {
            before("Test Server should be started", async function () {
                const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
                const serverResponse = await testHelper.pingTestServer(port);
                assert.equal(testServerStarted, true);
                assert.equal(serverResponse["status"], 200);
            }),
            describe("Start dev-server and sideload application", function () {
                it(`Sideload should have completed for ${host} and dev-server should have started`, async function () {
                    this.timeout(0);
                    const startDevServer = await testHelper.startDevServer();
                    const sideloadApplication = await testHelper.sideloadDesktopApp(host, manifestPath);
                    assert.equal(startDevServer, true);
                    assert.equal(sideloadApplication, true);
                });
            });
            describe("Get test results for taskpane project", function () {
                it("Validate expected result count", async function () {
                    this.timeout(0);
                    testValues = await testServer.getTestResults();
                    assert.equal(testValues.length > 0, true);
                });
                it("Validate expected result name", async function () {
                    assert.equal(testValues[0].Name, resultName);
                });
                it("Validate expected result", async function () {
                    assert.equal(testValues[0].Value, resultValue);
                });
            });
            after("Teardown test environment", async function () {
                const stopTestServer = await testServer.stopTestServer();
                assert.equal(stopTestServer, true);
                const testEnvironmentTornDown = await testHelper.teardownTestEnvironment(host)
                assert.equal(testEnvironmentTornDown, true);
            });
        });
    });
}





