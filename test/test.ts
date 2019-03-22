import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
const testJsonFile: string = `${__dirname}/testData.json`;
const testJsonData = JSON.parse(fs.readFileSync(testJsonFile).toString());
import * as testHelper from "office-addin-test-helpers"
import * as testServerInfra from "office-addin-test-server";
const port: number = 8080;
const testServer = new testServerInfra.TestServer(port);
let testValues : any = [];

Object.keys(testJsonData.hosts).forEach(function (host) {
    const expectedResult: string = testJsonData.hosts[host].expectedResult;

    describe("Setup test environment", function () {
        describe("Start sideload, start dev-server, and start test-server", function () {
            it(`Sideload should have completed for ${host} and dev-server should have started`, async function () {
                this.timeout(0);
                const startDevServer = await testHelper.startDevServer();
                const sideloadApplication = await testHelper.sideloadDesktopApp(host, "test/test-manifest.xml");
                assert.equal(startDevServer, true);
                assert.equal(sideloadApplication, true);
            });
            it(`Test server should have started and ${host} should have pinged the server`, async function () {
                this.timeout(0);
                const testServerStarted = await testServer.startTestServer();
                assert.equal(testServerStarted, true);
            });
        });
    });

    describe("Test Taskpane Project", function () {
        describe("Get test results for taskpane project", function () {
            it("Validate expected result count", async function () {
                this.timeout(0);
                testValues = await testServer.getTestResults();
                assert.equal(testValues.length > 0, true);
            });
            it("Validate expcted result", async function () {
                assert.equal(testValues[0].Value, expectedResult);
            });
        });
    });

    describe("Teardown test environment", function () {
        describe(`Kill ${host} and the test server`, function () {
            it(`should close ${host} and stop the test server`, async function () {
                this.timeout(10000);
                const stopTestServer = await testServer.stopTestServer();
                assert.equal(stopTestServer, true);
                const testEnvironmentTornDown = await testHelper.teardownTestEnvironment(host)
                assert.equal(testEnvironmentTornDown, true);
            });
        });
    })
});



