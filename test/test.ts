import * as testInfra from "office-addin-test-server";
import * as assert from "assert";
import * as mocha from "mocha";
const yellowColor: string = "#FFFF00";
const port: number = 8080;
import TestServer from "../node_modules/office-addin-test-server/lib/testServer.js";
const testServer = new TestServer(port);
let testValues : any = [];
const promiseSetupTestEnvironment = testInfra.setupTestEnvironment("excel", port);
const promiseStartTestServer = testServer.startTestServer();
const promiseGetTestResults = testServer.getTestResults();

describe("Setup test environment", function () {
    describe("Start sideload, start dev-server, and start test-server", function () {
        it("Sideload should have completed and dev-server should have started", async function () {
            const setupTestEnvironmentSucceeded = await promiseSetupTestEnvironment;
            assert.equal(setupTestEnvironmentSucceeded, true);
        });
        it("Test server should have started and Excel should have pinged the server", async function () {
            const testServerStarted = await promiseStartTestServer;
                assert.equal(testServerStarted, true);
        });
    });
});

describe("Test Taskpane Project", function () {
    describe("Get test results for taskpane project", function () {
        it("should get results from the taskpane application", async function () {
            testValues = await promiseGetTestResults;
            // Expecting one result
            assert.equal(testValues.length, 1);
        });
        it("Cell font color should have changes to red", async function () {
            assert.equal(yellowColor, testValues[0].Value);
        });
    });
});

describe("Teardown test environment", function () {
    describe("Kill Excel and the test server", function () {
        it("should close Excel and stop the test server", async function () {
            await testInfra.teardownTestEnvironment(process.platform == 'win32' ? "EXCEL" : "Excel");
        });
    });
})