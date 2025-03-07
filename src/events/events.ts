import { runOutlook } from "../shared/outlook";

export function onMessageCompose(event: Office.MailboxEvent) {
  runOutlook("Message compose event");
  event.completed();
}

Office.actions.associate("onMsgCompose", onMessageCompose);
