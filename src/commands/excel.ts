/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runExcel } from "../shared/excel";

/* global Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    Office.actions.associate("action", runExcelCommand);
  }
});

function runExcelCommand(event: Office.AddinCommands.Event): void {
  runExcel();

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}
