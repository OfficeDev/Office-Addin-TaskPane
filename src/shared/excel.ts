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
