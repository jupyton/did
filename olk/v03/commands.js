const MY_NAME = 'v02 - 010';


const ssnRegex = /\b(\d{3}-\d{2}-\d{4}|\d{9})\b/g;
const nricRegex = /[SFTG]\d{7}[A-Z]/gm;
const creditcardRegex = /(\d{4}[-]){3}\d{4}|\d{16}/gm;



const nricRedacted = 'X0000000X';
const creditcardRedacted = 'xxxx-xxxx-xxxx-xxxx';


Office.onReady((info) => {
  console.info('[v03] Commands.js::onReady()');
});

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// Factories
const makePromiseSetSubject = (mailItem, newSubject) => {
  return new Promise((resolve, reject) => {
    console.info("[v03] trying to set subject to [" + newSubject + "]");
    mailItem.subject.setAsync(newSubject, { coercionType: Office.CoercionType.subjectHtml }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v03] set subject OK : [" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.info("[v03] set subject BAD : " + setAsyncResult.error.message + "]");
        reject(setAsyncResult.error.message);
      }
    });
  });
};

const makePromiseSetBody = (mailItem, newBody) => {
  return new Promise((resolve, reject) => {
    console.info("[v03] trying to set body to [" + newBody + "]");
    mailItem.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v03] set body OK : [" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.info("[v03] set body BAD : [" + setAsyncResult.error.message + "]");
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

  let senderName = 'unknown';
  let senderEmail = 'unknown';
  if (item.sender) {
    senderName = item.sender.displayName;
    senderEmail = item.sender.emailAddress;
  }

  console.log(`[v03] Sender Display Name: ${senderName}`);
  console.log(`[v03] Sender Email Address: ${senderEmail}`);


  const getSubjectPromise = new Promise((resolve, reject) => {
    console.info("[v03] trying to get subject");
    item.subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v03] fetch subject OK");
        resolve(asyncResult.value);
      } else {
        console.info("[v03] fetch subject BAD");
        reject(asyncResult.error.message);
      }
    });
  });

  const getBodyPromise = new Promise((resolve, reject) => {
    console.info("[v03] trying to get body");
    item.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v03] fetch body OK");
        resolve(asyncResult.value);
      } else {
        console.info("[v03] fetch body BAD");
        reject(asyncResult.error.message);
      }
    });
  });




  Promise.all([getSubjectPromise, getBodyPromise]).then(([subjectHtml, bodyHtml]) => {
    subjectHtml = subjectHtml + "";
    subjectHtml = subjectHtml.trim();
    bodyHtml = bodyHtml + "";
    bodyHtml = bodyHtml.trim();

    console.info("[v03] SUBJECT --->");
    console.info("[v03] " + subjectHtml);
    console.info("[v03] <--- SUBJECT");

    console.info("[v03] BODY --->");
    console.info("[v03] " + bodyHtml);
    console.info("[v03] <--- BODY");

    sanitizedSubjectHtml = subjectHtml.replace(nricRegex, nricRedacted);
    sanitizedSubjectHtml = sanitizedSubjectHtml.replace(creditcardRegex, creditcardRedacted);

    sanitizedBodyHtml = bodyHtml.replace(nricRegex, nricRedacted);
    sanitizedBodyHtml = sanitizedBodyHtml.replace(creditcardRegex, creditcardRedacted);

    console.info("[v03] Sanitized SUBJECT --->");
    console.info("[v03] " + sanitizedSubjectHtml);
    console.info("[v03] <--- Sanitized SUBJECT");

    console.info("[v03] Sanitized BODY --->");
    console.info("[v03] " + sanitizedBodyHtml);
    console.info("[v03] <--- Sanitized BODY");

    let promiseSetSubject = makePromiseSetSubject(item, sanitizedSubjectHtml);
    let promiseSetBody = makePromiseSetBody(item, sanitizedBodyHtml);

    Promise.all([promiseSetSubject, promiseSetBody]).then(() => {
      console.info("[v03] successfully set redacted SUBJECT / BODY:");


      // POST to LOG Server
      fetch('https://jsonplaceholder.typicode.com/comments', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          postId: 89,
          name: `whose name [${MY_NAME}]`,
          email: "whose@name.world",
          body: new Date().toISOString()
        }),
      })
        .then(response => {
          console.info(`[v03] in fetch() response, STATUS=[${response.statusText}]`);

          if (!response.ok) {
            console.error(`[v03] POST to LOG - API failed, STATUS=[${response.statusText}]`);
            event.completed({ allowEvent: false, errorMessage: "Send LOG failed on API" });
          }

          console.info(`[v03] POST to LOG - API OK, STATUS=[${response.statusText}]`);

          return response.json();
        })
        .then(data => {
          console.info(`[v03] POST to LOG - API OK, DATA=[${data}]`);
          console.info(`[v03] POST to LOG - API OK, DATA JSON=[${JSON.stringify(data)}]`);

          Office.context.mailbox.item.notificationMessages.replaceAsync('redacter', {
            type: 'errorMessage',
            message: `Everything OK, but still don't let you send [${MY_NAME}]`
          }, function (result) {
          });

          event.completed({
            allowEvent: false,
            errorMessage: "Everything OK, but still don't let you send"
          });
        })
        .catch(error => {
          console.error("[v03] POST to LOG - NETWORK failed :", error);
          event.completed({ allowEvent: false, errorMessage: "Send LOG failed on Network" });
        });

      // POST to LOG Server




      
    }).catch((error) => {
      console.error("[v03] An error occurred while setting item data:", error);
      event.completed({ allowEvent: false, errorMessage: "Not able to set SUBJECT, BODY!!!" });
    });
  }).catch((error) => {
    console.error("[v03] An error occurred while fetching item data:", error);
    event.completed({ allowEvent: false, errorMessage: "Not able to retrieve SUBJECT, BODY!!!" });
  });






  console.info("[v02 Commands.js::onMessageSendHandler(100)] Exit!");
}

