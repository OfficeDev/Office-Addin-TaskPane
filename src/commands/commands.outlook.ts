import { setNotificationInOutlook } from "./outlook";

/* global Office */

// Register the add-in commands with the Office host application.
Office.onReady(async () => {
  Office.actions.associate("action", setNotificationInOutlook);
});
