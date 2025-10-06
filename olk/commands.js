Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

function onMessageSendHandler(event) {
  console.log("[Commands.js::onMessageSendHandler] The OnMessageSend event was triggered!");

  event.completed({ allowEvent: false });
}
