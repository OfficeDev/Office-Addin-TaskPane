import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";

/* global describe, global, it, require */

// office-addin-mock needs to be able to handle collections like "Slides" and "Shapes" before we can fully verify load and sync behavior.
// For now, we're using and not completely mocked object to verify the general flow.

const slide = {
  shapes: {
    addTextBox: function (text: string, options: any) {
      const shape = {
        name: "",
        textFrame: {
          textRange: {
            text: text,
          },
        },
      };
      this.items.push(shape);
      return shape;
    },
    items: [],
  },
};

const PowerPointMockData = {
  context: {
    presentation: {
      slides: {
        getItemAt(index: number) {
          return slide;
        },
      },
    },
  },
  run: async function (callback) {
    await callback(this.context);
  },
};

const OfficeMockData = {
  onReady: async function () {},
};

describe("PowerPoint", function () {
  it("Run", async function () {
    
    const pptMock: OfficeMockObject = new OfficeMockObject(PowerPointMockData); // Mocking the host specific namespace
    global.PowerPoint = pptMock as any;
    global.Office = new OfficeMockObject(OfficeMockData) as any; // Mocking the common office-js namespace

    const { run } = require("../../src/taskpane/powerpoint");
    await run();

    // Check that a text box was added with the correct text
    assert.strictEqual(slide.shapes.items[0].textFrame.textRange.text, "Hello World!");
  });
});
