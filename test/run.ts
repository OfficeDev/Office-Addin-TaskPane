import * as assert from "assert";
import { OfficeJSMock } from "./mock_utils";
import { run } from "../src/test-file";
const JsonData = require("./run.json");

/* global describe, global, it */

describe(`Run`, function () {
  it("Using json", async function () {
    const excelMock = new OfficeJSMock("excel") as any;
    excelMock.populate(JsonData);

    excelMock.context.workbook.addMockFunction("getSelectedRange", () => excelMock.context.workbook.range);
    excelMock.addMockFunction("run", async function(callback) {
      await callback(excelMock.context);
    });
  
    global.Excel = excelMock;
  
    await run();
    assert.strictEqual(excelMock.context.workbook.range.format.fill.color, "yellow");
  });
  it("Basic test", async function () {
    const excelMock = new OfficeJSMock("excel") as any;

    excelMock.addMockObject("context");
    excelMock.context.addMockObject("workbook");
    excelMock.context.workbook.addMockObject("range");
    excelMock.context.workbook.addMockFunction("getSelectedRange", () => excelMock.context.workbook.range);
    excelMock.context.workbook.range.setMock("address", "G4");
    excelMock.context.workbook.range.addMockObject("format");
    excelMock.context.workbook.range.format.addMockObject("fill");
    excelMock.addMockFunction("run", async function(callback) {
      await callback(excelMock.context);
    });
  
    global.Excel = excelMock;
  
    await run();
    assert.strictEqual(excelMock.context.workbook.range.format.fill.color, "yellow");
  });
});