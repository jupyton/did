const MY_NAME = 'v02 - 005';


const ssnRegex = /\b(\d{3}-\d{2}-\d{4}|\d{9})\b/g;
const nricRegex = /[SFTG]\d{7}[A-Z]/gm;
const creditcardRegex = /(\d{4}[-]){3}\d{4}|\d{16}/gm;



const nricRedacted = 'X0000000X';
const creditcardRedacted = 'xxxx-xxxx-xxxx-xxxx';


Office.onReady((info) => {
  console.info('[v02] Commands.js::onReady()');
});

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// Factories
const makePromiseSetSubject = (mailItem, newSubject) => {
  return new Promise((resolve, reject) => {
    console.info("[v02] trying to set subject to [" + newSubject + "]");
    mailItem.subject.setAsync(newSubject, { coercionType: Office.CoercionType.subjectHtml }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v02] set subject OK : [" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.info("[v02] set subject BAD : " + setAsyncResult.error.message + "]");
        reject(setAsyncResult.error.message);
      }
    });
  });
};

const makePromiseSetBody = (mailItem, newBody) => {
  return new Promise((resolve, reject) => {
    console.info("[v02] trying to set body to [" + newBody + "]");
    mailItem.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v02] set body OK : [" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.info("[v02] set body BAD : [" + setAsyncResult.error.message + "]");
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
  console.info("[v02 Commands.js::onMessageSendHandler()] Received OnMessageSend event!");

  Office.context.mailbox.item.notificationMessages.replaceAsync('redacter', {
    type: 'errorMessage',
    message: "Argentra notificationMessages " + MY_NAME
  }, function (result) {
  });


  const item = Office.context.mailbox.item;
  let sanitizedSubjectHtml = "";
  let sanitizedBodyHtml = "";

  const getSubjectPromise = new Promise((resolve, reject) => {
    console.info("[v02] trying to get subject");
    item.subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v02] fetch subject OK");
        resolve(asyncResult.value);
      } else {
        console.info("[v02] fetch subject BAD");
        reject(asyncResult.error.message);
      }
    });
  });

  const getBodyPromise = new Promise((resolve, reject) => {
    console.info("[v02] trying to get body");
    item.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v02] fetch body OK");
        resolve(asyncResult.value);
      } else {
        console.info("[v02] fetch body BAD");
        reject(asyncResult.error.message);
      }
    });
  });




  Promise.all([getSubjectPromise, getBodyPromise]).then(([subjectHtml, bodyHtml]) => {
    subjectHtml = subjectHtml + "";
    subjectHtml = subjectHtml.trim();
    bodyHtml = bodyHtml + "";
    bodyHtml = bodyHtml.trim();

    console.info("[v02] SUBJECT --->");
    console.info("[v02] " + subjectHtml);
    console.info("[v02] <--- SUBJECT");

    console.info("[v02] BODY --->");
    console.info("[v02] " + bodyHtml);
    console.info("[v02] <--- BODY");

    sanitizedSubjectHtml = subjectHtml.replace(nricRegex, nricRedacted);
    sanitizedSubjectHtml = sanitizedSubjectHtml.replace(creditcardRegex, creditcardRedacted);

    sanitizedBodyHtml = bodyHtml.replace(nricRegex, nricRedacted);
    sanitizedBodyHtml = sanitizedBodyHtml.replace(creditcardRegex, creditcardRedacted);

    console.info("[v02] Sanitized SUBJECT --->");
    console.info("[v02] " + sanitizedSubjectHtml);
    console.info("[v02] <--- Sanitized SUBJECT");

    console.info("[v02] Sanitized BODY --->");
    console.info("[v02] " + sanitizedBodyHtml);
    console.info("[v02] <--- Sanitized BODY");

    let promiseSetSubject = makePromiseSetSubject(item, sanitizedSubjectHtml);
    let promiseSetBody = makePromiseSetBody(item, sanitizedBodyHtml);

    Promise.all([promiseSetSubject, promiseSetBody]).then(() => {
      console.info("[v02] successfully set redacted SUBJECT / BODY:");


      // POST to LOG Server
      fetch('https://jsonplaceholder.typicode.com/comments', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          postId: 89,
          name: 'whose name',
          email: "whose@names.world",
          body: new Date().toISOString()
        }),
      })
        .then(response => {
          console.info(`[v02] in fetch() response, STATUS=[${response.statusText}]`);

          if (!response.ok) {
            console.error(`[v02] POST to LOG - API failed, STATUS=[${response.statusText}]`);
            event.completed({ allowEvent: false, errorMessage: "Send LOG failed on API" });
          }

          console.info(`[v02] POST to LOG - API OK, STATUS=[${response.statusText}]`);

          return response.json();
        })
        .then(dataBody => {
          console.info(`[v02] POST to LOG - API OK, BODY=[${dataBody}]`);
          event.completed({
            allowEvent: false,
            errorMessage: "Everything OK, but still don't let you send"
          });
        })
        .catch(error => {
          console.error("[v02] POST to LOG - NETWORK failed :", error);
          event.completed({ allowEvent: false, errorMessage: "Send LOG failed on Network" });
        });

      // POST to LOG Server




      
    }).catch((error) => {
      console.error("[v02] An error occurred while setting item data:", error);
      event.completed({ allowEvent: false, errorMessage: "Not able to set SUBJECT, BODY!!!" });
    });
  }).catch((error) => {
    console.error("[v02] An error occurred while fetching item data:", error);
    event.completed({ allowEvent: false, errorMessage: "Not able to retrieve SUBJECT, BODY!!!" });
  });






  console.info("[v02 Commands.js::onMessageSendHandler(100)] Exit!");
}

