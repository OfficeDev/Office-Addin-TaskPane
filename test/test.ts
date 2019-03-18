import * as assert from "assert";
import * as mocha from "mocha";
import * as testHelper from "office-addin-test-helpers"
import * as testServerInfra from "office-addin-test-server";
const yellowColor: string = "#FFFF00";
const port: number = 8080;
const testServer = new testServerInfra.TestServer(port);
let testValues : any = [];

describe("Setup test environment", function () {
    describe("Start sideload, start dev-server, and start test-server", function () {
        it("Sideload should have completed and dev-server should have started", async function () {
            this.timeout(0);
            const startDevServer = await testHelper.startDevServer();
            const sideloadApplication = await testHelper.sideloadDesktopApp("excel", "test/test-manifest.xml");
            assert.equal(startDevServer, true);
            assert.equal(sideloadApplication, true);
        });
        it("Test server should have started and Excel should have pinged the server", async function () {
            this.timeout(0);
            const testServerStarted = await testServer.startTestServer();
            assert.equal(testServerStarted, true);
        });
    });
 });

describe("Test Taskpane Project", function () {
    describe("Get test results for taskpane project", function () {
        it("should get results from the taskpane application", async function () {
            this.timeout(0);
            testValues = await testServer.getTestResults();
            assert.equal(testValues.length, 1);
        });
        it("Cell fill color should be yellow", async function () {
            assert.equal(yellowColor, testValues[0].Value);
        });
    });
});

describe("Teardown test environment", function () {
    describe("Kill Excel and the test server", function () {
        it("should close Excel and stop the test server", async function () {
            this.timeout(10000);
            const stopTestServer = await testServer.stopTestServer();
            assert.equal(stopTestServer, true);
            await testHelper.teardownTestEnvironment(process.platform == 'win32' ? "EXCEL" : "Excel");
        });
    });
})