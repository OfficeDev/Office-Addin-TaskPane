/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runOutlook } from "../shared/outlook";

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    Office.actions.associate("action", runOutlookCommand);
  }
});

function runOutlookCommand(event: Office.AddinCommands.Event): void {
  runOutlook("Clicked command button");

  // Be sure to indicate when the add-in command function is complete.
  if (event) {
    event.completed();
  }
}
