/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

export async function runPowerPoint(): Promise<void> {
  /**
   * Insert your PowerPoint code here
   */

  const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };
  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}

export async function insertTextInPowerPoint(text: string): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const slide: PowerPoint.Slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox: PowerPoint.Shape = slide.shapes.addTextBox(text);
      textBox.fill.setSolidColor("white");
      textBox.lineFormat.color = "black";
      textBox.lineFormat.weight = 1;
      textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
