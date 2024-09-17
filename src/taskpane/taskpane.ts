/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "./excel";
import "./outlook";
import "./powerpoint";
import "./word";

/* global Office, document */

Office.onReady(async () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
});
