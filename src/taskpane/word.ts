/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runWord } from "../shared/word";

/* global document Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("run").onclick = runWord;
  }
});
