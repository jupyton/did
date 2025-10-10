const ssnRegex = /\b(\d{3}-\d{2}-\d{4}|\d{9})\b/g;
const nricRegex = /[SFTG]\d{7}[A-Z]/gm;
const creditcardRegex = /(\d{4}[-]){3}\d{4}|\d{16}/gm;



const nricRedacted = 'X0000000X';
const creditcardRedacted = 'xxxx-xxxx-xxxx-xxxx';


Office.onReady((info) => {
  console.info('Commands.js::onReady()');
});

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// Factories
const makePromiseSetSubject = (mailItem, newSubject) => {
  return new Promise((resolve, reject) => {
    console.log("[ARG] trying to set subject to [" + newSubject + "]");
    mailItem.subject.setAsync(newSubject, { coercionType: Office.CoercionType.subjectHtml }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[ARG] set subject OK" + setAsyncResult.value);
        resolve(setAsyncResult.value);
      } else {
        console.log("[ARG] set subject BAD" + setAsyncResult.error.message);
        reject(setAsyncResult.error.message);
      }
    });
  });
};

const makePromiseSetBody = (mailItem, newBody) => {
  return new Promise((resolve, reject) => {
    console.log("[ARG] trying to set body to [" + newBody + "]");
    mailItem.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[ARG] set body OK" + setAsyncResult.value);
        resolve(setAsyncResult.value);
      } else {
        console.log("[ARG] set body BAD" + setAsyncResult.error.message);
        reject(setAsyncResult.error.message);
      }
    });
  });
};


// event handler
function onMessageSendHandler(event) {
  console.info("[Commands.js::onMessageSendHandler()] Received OnMessageSend event!");

  // ======== TEST - show a NOTIFICATION ========
  Office.context.mailbox.item.notificationMessages.replaceAsync('github-error', {
    type: 'errorMessage',
    message: 'is there an error? block first!'
  }, function(result){
  });
  // ======== TEST - show a NOTIFICATION ========


  // ======== TEST - show a DIALOG ========
  const url = 'https://jupyton.github.io/did/olk/dialog.html?warn=1'; // new URI('dialog.html?warn=1').absoluteTo(window.location).toString();
  console.info("[Commands.js::onMessageSendHandler()] url=[" + url + "]");
  const dialogOptions = { width: 20, height: 40, displayInIframe: true };

  Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
    settingsDialog = result.value;
    console.info("[Commands.js::onMessageSendHandler()] settingsDialog=[" + settingsDialog + "]");
    // settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
    // settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
  });
  // ======== TEST - show a DIALOG ========

  if (1<0) {

  const item = Office.context.mailbox.item;
  let sanitizedSubjectHtml = "";
  let sanitizedBodyHtml = "";

  const getSubjectPromise = new Promise((resolve, reject) => {
    console.log("[ARG] trying to get subject");
    item.subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[ARG] fetch subject OK");
        resolve(asyncResult.value);
      } else {
        console.log("[ARG] fetch subject BAD");
        reject(asyncResult.error.message);
      }
    });
  });

  const getBodyPromise = new Promise((resolve, reject) => {
    console.log("[ARG] trying to get body");
    item.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[ARG] fetch body OK");
        resolve(asyncResult.value);
      } else {
        console.log("[ARG] fetch body BAD");
        reject(asyncResult.error.message);
      }
    });
  });



    
    Promise.all([getSubjectPromise, getBodyPromise]).then(([subjectHtml, bodyHtml]) => {
        subjectHtml = subjectHtml + "";
        subjectHtml = subjectHtml.trim();
        bodyHtml = bodyHtml + "";
        bodyHtml = bodyHtml.trim();

        console.log("[ARG] SUBJECT --->");
        console.log("[ARG] " + subjectHtml);
        console.log("[ARG] <--- SUBJECT");

        console.log("[ARG] BODY --->");
        console.log("[ARG] " + bodyHtml);
        console.log("[ARG] <--- BODY");

        sanitizedSubjectHtml = subjectHtml.replace(nricRegex, nricRedacted);
        sanitizedSubjectHtml = sanitizedSubjectHtml.replace(creditcardRegex, creditcardRedacted);

        sanitizedBodyHtml = bodyHtml.replace(nricRegex, nricRedacted);
        sanitizedBodyHtml = sanitizedBodyHtml.replace(creditcardRegex, creditcardRedacted);

        console.log("[ARG] Sanitized SUBJECT --->");
        console.log("[ARG] " + sanitizedSubjectHtml);
        console.log("[ARG] <--- Sanitized SUBJECT");

        console.log("[ARG] Sanitized BODY --->");
        console.log("[ARG] " + sanitizedBodyHtml);
        console.log("[ARG] <--- Sanitized BODY");

        let promiseSetSubject = makePromiseSetSubject(item, sanitizedSubjectHtml);
        let promiseSetBody = makePromiseSetBody(item, sanitizedBodyHtml);

        Promise.all([promiseSetSubject, promiseSetBody]).then(() => {
            console.info("[ARG] successfully set redacted SUBJECT / BODY:");
            event.completed({ allowEvent: false, errorMessage: "Everything is fine. But I just want to block!!!" });
        }).catch((error) => {
            console.error("[ARG] An error occurred while setting item data:", error);
            event.completed({ allowEvent: false, errorMessage: "Not able to set SUBJECT, BODY!!!" });
        });
    }).catch((error) => {
        console.error("[ARG] An error occurred while fetching item data:", error);
        event.completed({ allowEvent: false, errorMessage: "Not able to retrieve SUBJECT, BODY!!!" });
    });

  }
  

  

  console.info("[Commands.js::onMessageSendHandler()] Exit!");
}



