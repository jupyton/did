const MY_NAME = 'v08o - 001'; //version


Office.onReady((info) => {
  console.info(`[ARG-o] ${MY_NAME} Commands.js::onReady()`);
});

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);






function onMessageSendHandler(event) {
  console.info("[ARG-o] Commands.js::onMessageSendHandler(): Received OnMessageSend event!");



            event.completed({
              allowEvent: true,
              errorMessage: "Everything OK, but still don't let you send"
            });






  console.info(`[${MY_NAME}] Commands.js::onMessageSendHandler(100)] Exit!`);
}

