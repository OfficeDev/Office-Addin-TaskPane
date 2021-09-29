import * as assert from "assert";
import { OfficeMockObject } from "office-addin-mock";
import { run } from "../../src/taskpane/excel";

const MockData = {
  context: {
    workbook: {
      range: {
        address: "G4",
        format: {
          fill: {},
        },
      },
      getSelectedRange: function () {
        return this.range;
      },
    },
  },
};

/* global describe, global, it */

describe(`Excel`, function () {
  it("Run", async function () {
    const excelMock = new OfficeMockObject(MockData) as any;
    excelMock.addMockFunction("run", async function (callback) {
      await callback(excelMock.context);
    });
    global.Excel = excelMock;

    await run();

    assert.strictEqual(excelMock.context.workbook.range.format.fill.color, "yellow");
  });
});