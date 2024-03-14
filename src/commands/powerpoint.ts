/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
    // If needed, Office.js is ready to be called.
  });
  
  /**
   * Shows a notification when the add-in command is executed.
   * @param event
   */
  async function action(event: Office.AddinCommands.Event) {
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
  
  // Register the function with Office.
  Office.actions.associate("action", action);
  