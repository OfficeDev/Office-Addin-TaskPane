/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runOutlook } from "../shared/outlook";

/* global document Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("run").onclick = () => {
      runOutlook("Clicked the taskpane button");
    };
  }
});
