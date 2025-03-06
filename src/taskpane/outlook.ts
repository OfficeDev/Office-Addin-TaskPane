/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runOutlook } from "../shared/outlook";

/* global document Office */

Office.onReady((info: any) => {
  if (info.host === Office.HostType.Outlook) {
    const runButton = document.getElementById("run");
    if (runButton) {
      runButton.onclick = () => {
        runOutlook("Clicked the Task Pane button");
      };
    }
  }
});
