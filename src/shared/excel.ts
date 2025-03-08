/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console Excel */

export async function runExcel(): Promise<void> {
  await Excel.run(async (context: Excel.RequestContext) => {
    /**
     * Insert your Excel code here
     */

    const range: Excel.Range = context.workbook.getSelectedRange();
    range.load("address");
    range.format.fill.color = "yellow";
    await context.sync();
    console.log(`The range address was ${range.address}.`);
  });
}

export async function insertTextInExcel(text: string): Promise<void> {
  // Write text to the top left cell.
  try {
    await Excel.run(async (context) => {
      const sheet: Excel.Worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range: Excel.Range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
