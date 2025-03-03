/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "./excel";
import "./outlook";
import "./powerpoint";
import "./word";

/* global Office document */

Office.onReady(async () => {
  const sideloadMsg = document.getElementById("sideload-msg");
  if (sideloadMsg) {
    sideloadMsg.style.display = "none";
  }
  
  const appBody = document.getElementById("app-body");
  if (appBody) {
    appBody.style.display = "flex";
  }
});
