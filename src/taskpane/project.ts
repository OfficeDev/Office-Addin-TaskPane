/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { runProject } from "../shared/project";

/* global document Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Project) {
    document.getElementById("run").onclick = runProject;
  }
});
