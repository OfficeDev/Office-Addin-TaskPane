import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";

/* global describe, global, it, require */

const PowerPointMockData = {
  context: {
    presentation: {
      // office-addin-mock needs to be able to handle collections like "Slides" and "Shapes" before this can be used including appropriate load and sync behavior.
      // For now, this is a prototype of structure needed to support the test in the future.
      slides: {
        items: [
          {
            shapes: {
              items: [],
              addTextBox(text: string, options: any) {
                const textBox = new OfficeMockObject({
                  name: "",
                  textFrame: {
                    textRange: {
                      text: text,
                    },
                  },
                });
                this.items.push(textBox);
                return textBox;
              },
            },
          },
        ],
        getItemAt(index: number) {
          return this.items[index];
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
    const textBox = pptMock.context.presentation.slides.items[0].shapes.items[0];
    assert.strictEqual(textBox.textFrame.textRange.text, "Hello World!");
  });
});
