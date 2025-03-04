/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runPowerPoint } from "../shared/powerpoint";

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    Office.actions.associate("action", runPowerPointCommand);
  }
});

function runPowerPointCommand(event: Office.AddinCommands.Event): void {
  runPowerPoint();

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
