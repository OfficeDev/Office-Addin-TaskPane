/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runExcel } from "../shared/excel";

/* global document Office */

Office.onReady((info: any) => {
  if (info.host === Office.HostType.Excel) {
    const runButton = document.getElementById("run");
    if (runButton) {
      runButton.onclick = runExcel;
    }
  }
});
