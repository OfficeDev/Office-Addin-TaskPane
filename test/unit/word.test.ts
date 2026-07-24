import * as assert from "assert";
import "mocha";
import { OfficeMockObject } from "office-addin-mock";

/* global describe, global, it, require, Word */

type MockParagraph = {
  font: {
    color?: string;
  };
  insertLocation?: string;
  text: string;
};

const wordMockContext = {
  document: {
    body: {
      paragraph: {
        font: {},
        text: "",
      } as MockParagraph,
      insertParagraph: function (paragraphText: string, insertLocation: Word.InsertLocation): MockParagraph {
        this.paragraph.text = paragraphText;
        this.paragraph.insertLocation = insertLocation;
        return this.paragraph;
      },
    },
  },
};

const WordMockData = {
  context: wordMockContext,
  InsertLocation: {
    end: "End",
  },
  run: async function (callback: (context: any) => Promise<void> | void) {
    await callback(this.context);
  },
};

const OfficeMockData = {
  onReady: async function () {},
};

describe("Word", function () {
  it("Run", async function () {
    const wordMock: OfficeMockObject = new OfficeMockObject(WordMockData); // Mocking the host specific namespace
    (global as any).Word = wordMock;
    (global as any).Office = new OfficeMockObject(OfficeMockData); // Mocking the common office-js namespace

    const { runWord } = require("../../src/taskpane/word");
    await runWord();

    assert.strictEqual(wordMock.context.document.body.paragraph.font.color, "blue");
  });
});
