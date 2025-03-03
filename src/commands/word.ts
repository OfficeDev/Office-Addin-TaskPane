/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runWord } from "../shared/word";

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    Office.actions.associate("action", runWordCommand);
  }
});

function runWordCommand(event: Office.AddinCommands.Event): void {
  runWord();

  // Be sure to indicate when the add-in command function is complete.
  if (event) {
    event.completed();
  }
}
