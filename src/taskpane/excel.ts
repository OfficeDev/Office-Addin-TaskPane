/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import * as usageWeb from "office-addin-usage-data-web";
declare var __USAGEDATAENABLED__: boolean;
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});
let usageDataInstance: usageWeb.OfficeAddinUsageData;
if(__USAGEDATAENABLED__) {
  var options = {
    groupName: usageWeb.groupName,
    projectName: "excel template test",
    instrumentationKey: usageWeb.instrumentationKeyForOfficeAddinCLITools,
    promptQuestion: "-----------------------------------------\nDo you want to opt-in for usage data?[y/n]\n-----------------------------------------",
    raisePrompt: false,
    usageDataLevel: usageWeb.UsageDataLevel.on,
    method: usageWeb.UsageDataReportingMethod.applicationInsights,
    isForTesting: true
  }
  usageDataInstance = new usageWeb.OfficeAddinUsageData(options);
}
export async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
      if(__USAGEDATAENABLED__) {
        if (usageDataInstance !== undefined) {
          console.log("telemetry here");
          usageDataInstance.reportEvent("rangeSelection", new Object(range.address));
        }
      }
    });
  } catch (error) {
    console.error(error);
  }
}
