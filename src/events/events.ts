import { runOutlook } from "../shared/outlook";

export function onMessageComposeHandler(event: Office.MailboxEvent) {
  runOutlook("Message compose event");
  event.completed();
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
