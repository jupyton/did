const VERSION = 'v001.0003';


Office.onReady((info) => {
  console.info(`Commands.js::onReady(${VERSION})`);
});

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// Factories
const makePromiseSetSubject = (mailItem, newSubject) => {
  return new Promise((resolve, reject) => {
    console.log("[ARG] trying to set subject to [" + newSubject + "]");
    mailItem.subject.setAsync(newSubject, { coercionType: Office.CoercionType.subjectHtml }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[ARG] set subject OK : [" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.log("[ARG] set subject BAD : " + setAsyncResult.error.message + "]");
        reject(setAsyncResult.error.message);
      }
    });
  });
};

const makePromiseSetBody = (mailItem, newBody) => {
  return new Promise((resolve, reject) => {
    console.log("[ARG] trying to set body to (((" + newBody + ")))");
    mailItem.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log("[ARG] set body OK : [" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.log("[ARG] set body BAD : [" + setAsyncResult.error.message + "]");
        reject(setAsyncResult.error.message);
      }
    });
  });
};


// event handler
function redactMessageHandler() {
  console.info("[Commands.js::redactMessageHandler()] being called!");
}


function onMessageSendHandler(event) {
  console.info(`[onMessageSendHandler(0000)] Received OnMessageSend event!`);



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



  Promise.all([getSubjectPromise]).then(([subjectHtml]) => {
    subjectHtml = subjectHtml + "";
    subjectHtml = subjectHtml.trim();

    console.log("[ARG] SUBJECT --->");
    console.log("[ARG] " + subjectHtml);
    console.log("[ARG] <--- SUBJECT");


    sanitizedBodyHtml = "Copied Subject to Body as (((" + subjectHtml + ")))";

    console.log("[ARG] Sanitized BODY --->");
    console.log("[ARG] " + sanitizedBodyHtml);
    console.log("[ARG] <--- Sanitized BODY");

    let promiseSetBody = makePromiseSetBody(item, sanitizedBodyHtml);

    Promise.all([promiseSetBody]).then(() => {
      console.info("[ARG] successfully copied SUBJECT to BODY:");
      event.completed({
        allowEvent: false,
        errorMessage: "Everything OK, but still don't let you send"
      });
    }).catch((error) => {
      console.error("[ARG] An error occurred while setting item data:", error);
      event.completed({ allowEvent: false, errorMessage: "Not able to set SUBJECT, BODY!!!" });
    });
  }).catch((error) => {
    console.error("[ARG] An error occurred while fetching item data:", error);
    event.completed({ allowEvent: false, errorMessage: "Not able to retrieve SUBJECT, BODY!!!" });
  });






  console.info("[onMessageSendHandler(9999)] Exit!");
}

