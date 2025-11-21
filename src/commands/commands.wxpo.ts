import { insertBlueParagraphInWord } from "./commands.word";
import { setRangeColorInExcel } from "./commands.excel";
import { insertTextInPowerPoint } from "./commands.powerpoint";
import { setNotificationInOutlook } from "./commands.outlook";

/* global Office */

// Register the add-in commands with the Office host application.
Office.onReady(async (info) => {
  switch (info.host) {
    case Office.HostType.Word:
      Office.actions.associate("action", insertBlueParagraphInWord);
      break;
    case Office.HostType.Excel:
      Office.actions.associate("action", setRangeColorInExcel);
      break;
    case Office.HostType.PowerPoint:
      Office.actions.associate("action", insertTextInPowerPoint);
      break;
    case Office.HostType.Outlook:
      Office.actions.associate("action", setNotificationInOutlook);
      break;
    default: {
      throw new Error(`${info.host} not supported.`);
    }
  }
});
