/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runOneNote } from "../shared/onenote";

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    Office.actions.associate("action", runOneNoteCommand);
  }
});

function runOneNoteCommand(event: Office.AddinCommands.Event) {
  runOneNote();

  // Be sure to indicate when the add-in command function is complete.
  if (event) {
    event.completed();
  }
}
