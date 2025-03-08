/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Word */

export async function runWord(): Promise<void> {
  return Word.run(async (context: Word.RequestContext) => {
    /**
     * Insert your Word code here
     */

    const paragraph: Word.Paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    paragraph.font.color = "blue";
    await context.sync();
  });
}

export async function insertTextInWord(text: string): Promise<void> {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body: Word.Body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
