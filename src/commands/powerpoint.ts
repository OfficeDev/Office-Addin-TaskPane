/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Register the function with Office.
    Office.actions.associate("action", actionPowerPoint);
  }
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
async function actionPowerPoint(event: Office.AddinCommands.Event) {
  try {
    await PowerPoint.run(async (context) => {
      const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };
      await Office.context.document.setSelectedDataAsync(" ", options);
      await Office.context.document.setSelectedDataAsync("Hello World!", options);
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}