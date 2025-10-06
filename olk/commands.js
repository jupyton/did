const ssnRegex = /\b(\d{3}-\d{2}-\d{4}|\d{9})\b/g; // /r[e3]views/g;
const nricRegex = /[SFTG]\d{7}[A-Z]/gm;
const creditcardRegex = /(\d{4}[-]){3}\d{4}|\d{16}/gm;


Office.onReady((info) => {
  console.info('Commands.js::onReady()');
});

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

function onMessageSendHandler(event) {
  console.info("[Commands.js::onMessageSendHandler()] Received OnMessageSend event!");


  // ---- SUBJECT ----
    Office.context.mailbox.item.subject.getAsync(
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const subjectHtml = asyncResult.value;
          console.info("[ARG] SUBJECT -->");
          console.info("[ARG] " + subjectHtml);
          console.info("[ARG] <-- SUBJECT");
          
          // More robust regex pattern for SSNs
          // const ssnRegex = /\b(\d{3}-\d{2}-\d{4}|\d{9})\b/g;
          
          // Replace matched SSNs with a masked version
          let sanitizedSubjectHtml = subjectHtml.replace(nricRegex, 'X0000000X');
          sanitizedSubjectHtml = sanitizedSubjectHtml.replace(creditcardRegex, 'xxxx-xxxx-xxxx-xxxx');
          console.info("Commands.js::onReady() CLEAN SUBJECT -->");
          console.info("Commands.js::onReady() " + sanitizedSubjectHtml);
          console.info("Commands.js::onReady() <-- CLEAN SUBJECT");

          // Set the sanitized HTML back into the email body
          Office.context.mailbox.item.subject.setAsync(sanitizedSubjectHtml, { coercionType: Office.CoercionType.subjectHtml }, function (setAsyncResult) {
            if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.info("Commands.js::onReady() S-1.001 SSNs have been redacted from the email subject.");
            } else {
              console.info("Commands.js::onReady() E-1.002 Failed to set the redacted email subject.");
            }
          });
          } else {
            console.info("[Commands.js::onMessageSendHandler()] E-1.001 Failed to get subject: " + asyncResult.error.message);
          }
      }
    );


  

  event.completed({ allowEvent: false });

  console.info("[Commands.js::onMessageSendHandler()] Exit!");
}


