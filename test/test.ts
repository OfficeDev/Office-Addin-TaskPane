import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import { AppType, startDebugging, stopDebugging} from "office-addin-debugging";
import * as path from "path";
import * as testHelper from "office-addin-test-helpers";
import * as testHelpers from "./src/test-helpers";
import * as testServerInfra from "office-addin-test-server";
const manifestPath = path.resolve(`${process.cwd()}/test/test-manifest.xml`);
const testJsonFile: string = path.resolve(`${process.cwd()}/test/src/test-data.json`);
const testJsonData = JSON.parse(fs.readFileSync(testJsonFile).toString());
const testServerPort: number = 4201;

Object.keys(testJsonData.hosts).forEach(function (host) {
    const testServer = new testServerInfra.TestServer(testServerPort);
    const resultName = testJsonData.hosts[host].resultName;
    const resultValue: string = testJsonData.hosts[host].resultValue;
    let testValues: any = [];

    describe(`Test ${host} Task Pane Project`, function () {
        before("Test Server should be started", async function () {
            const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
            const serverResponse = await testHelper.pingTestServer(testServerPort);
            assert.equal(testServerStarted, true);
            assert.equal(serverResponse["status"], 200);
        }),
            describe(`Start dev-server and sideload application: ${host}`, function () {
                it(`Sideload should have completed for ${host} and dev-server should have started`, async function () {
                    this.timeout(0);
                    const startDevServer = await testHelper.startDevServer();
                    assert.equal(startDevServer, true);

                    const sideloadCmd = `node ./node_modules/office-toolbox/app/office-toolbox.js sideload -m ${manifestPath} -a ${host}`;
                    await startDebugging(manifestPath, AppType.Desktop, undefined, undefined, undefined, undefined,
                        undefined, undefined, undefined, sideloadCmd);
                });
            });
        describe(`Get test results for ${host} taskpane project`, function () {
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
        after(`Teardown test environment and shutdown ${host}`, async function () {
            // Stop the test server
            const stopTestServer = await testServer.stopTestServer();
            assert.equal(stopTestServer, true);

            // Unregister the add-in
            const unregisterCmd = `node ./node_modules/office-toolbox/app/office-toolbox.js remove -m ${manifestPath} -a ${host}`;
            await stopDebugging(manifestPath, unregisterCmd);

            // Stop dev-server
            const testEnvironmentTornDown = await testHelper.teardownTestEnvironment(host);
            assert.equal(testEnvironmentTornDown, true);

            // Close desktop application for all apps but Excel, which has it's own closeWorkbook API
            if (host != 'Excel') {
                const applicationClosed = await testHelpers.closeDesktopApplication(host);
                assert.equal(applicationClosed, true);
            }
        });
    });
});



