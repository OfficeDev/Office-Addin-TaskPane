/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runPowerPoint } from "../shared/powerpoint";

/* global document Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    const runButton = document.getElementById("run");
    if (runButton) {
      runButton.onclick = runPowerPoint;
    }
  }
});
