/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runProject } from "../shared/project";

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Project) {
    Office.actions.associate("action", runProjectCommand);
  }
});

function runProjectCommand(event: Office.AddinCommands.Event) {
  runProject();

  // Be sure to indicate when the add-in command function is complete.
  if (event) {
    event.completed();
  }
}
