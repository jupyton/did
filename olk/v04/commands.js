const MY_NAME = 'v04 - 005';


const ssnRegex = /\b(\d{3}-\d{2}-\d{4}|\d{9})\b/g;
const nricRegex = /[SFTG]\d{7}[A-Z]/gm;
const creditcardRegex = /(\d{4}[-]){3}\d{4}|\d{16}/gm;



const nricRedacted = 'X0000000X';
const creditcardRedacted = 'xxxx-xxxx-xxxx-xxxx';


Office.onReady((info) => {
  console.info(`[v04] ${MY_NAME} Commands.js::onReady()`);
});

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// Factories
const makePromiseSetSubject = (mailItem, newSubject) => {
  return new Promise((resolve, reject) => {
    console.info("[v04] create PROMISE to SET SUBJECT to [" + newSubject + "]");
    mailItem.subject.setAsync(newSubject, { coercionType: Office.CoercionType.subjectHtml }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v04] SET SUBJECT OK : asyncResult-Value=[" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.info("[v04] SET SUBJECT BAD : " + setAsyncResult.error.message + "]");
        reject(setAsyncResult.error.message);
      }
    });
  });
};

const makePromiseSetBody = (mailItem, newBody) => {
  return new Promise((resolve, reject) => {
    console.info("[v04] create PROMISE to SET BODY to [" + newBody + "]");
    mailItem.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v04] SET BODY OK : asyncResult-Value=[" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.info("[v04] SET BODY BAD : [" + setAsyncResult.error.message + "]");
        reject(setAsyncResult.error.message);
      }
    });
  });
};


const makePromiseGetSubject = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info("[v04] create PROMISE to GET SUBJECT");
    mailItem.subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v04] GET SUBJECT OK");
        resolve(asyncResult.value);
      } else {
        console.info("[v04] GET SUBJECT BAD");
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetBody = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info("[v04] create PROMISE to GET BODY");
    mailItem.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v04] GET BODY OK");
        resolve(asyncResult.value);
      } else {
        console.info("[v04] GET BODY BAD");
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetTo = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info("[v04] create PROMISE to GET TO");
    mailItem.to.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v04] GET TO OK");
        resolve(asyncResult.value);
      } else {
        console.info("[v04] GET TO BAD");
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetFrom = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info("[v04] create PROMISE to GET FROM");
    mailItem.from.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[v04] GET FROM OK");
        resolve(asyncResult.value);
      } else {
        console.info("[v04] GET FROM BAD");
        reject(asyncResult.error.message);
      }
    });
  });
};


function onMessageSendHandler(event) {
  console.info("[v04] Commands.js::onMessageSendHandler(): Received OnMessageSend event!");


  //Office.context.mailbox.item.notificationMessages.replaceAsync('redacter', {
  //  type: 'errorMessage',
  //  message: "Argentra notificationMessages " + MY_NAME
  //}, function (result) {
  //});


  const item = Office.context.mailbox.item;
  let sanitizedSubjectHtml = "";
  let sanitizedBodyHtml = "";
  let mailboxEmail = "unknown";

  

  // ======== get MAILBOX
  const userProfile = Office.context.mailbox.userProfile;
  if (userProfile) {
    mailboxEmail = userProfile.emailAddress;
    if (mailboxEmail) {
      console.info(`[v04] SENDER MAILBOX=[${mailboxEmail}]`);
    } else {
      console.err("[v04] SENDER MAILBOX not available.");
    }
  } else {
    console.err("[v04] SENDER MAILBOX not available.");
  }




  const getToPromise = makePromiseGetTo(item);
  const getFromPromise = makePromiseGetFrom(item);
  const getSubjectPromise = makePromiseGetSubject(item);
  const getBodyPromise = makePromiseGetBody(item);

  Promise.all([getToPromise, getFromPromise, getSubjectPromise, getBodyPromise])
  .then(([to, from, subjectHtml, bodyHtml]) => {
    subjectHtml = subjectHtml + "";
    subjectHtml = subjectHtml.trim();
    bodyHtml = bodyHtml + "";
    bodyHtml = bodyHtml.trim();

    console.info("[v04] MAILBOX --->");
    console.info(`[v04]  - [${mailboxEmail}]`);
    console.info("[v04] <--- MAILBOX");

    console.info("[v04] FROM --->");
    console.info(`[v04]  - [${from.displayName} - ${from.emailAddress}]`);
    console.info("[v04] <--- FROM");

    console.info("[v04] TO --->");
    for (let i = 0; i < to.length; i++) {
      console.info(`[v04]  - [${to[i].displayName} - ${to[i].emailAddress}]`);
    }
    console.info("[v04] <--- TO");

    console.info("[v04] SUBJECT --->");
    console.info("[v04] " + subjectHtml);
    console.info("[v04] <--- SUBJECT");

    console.info("[v04] BODY --->");
    console.info("[v04] " + bodyHtml);
    console.info("[v04] <--- BODY");

    sanitizedSubjectHtml = subjectHtml.replace(nricRegex, nricRedacted);
    sanitizedSubjectHtml = sanitizedSubjectHtml.replace(creditcardRegex, creditcardRedacted);

    sanitizedBodyHtml = bodyHtml.replace(nricRegex, nricRedacted);
    sanitizedBodyHtml = sanitizedBodyHtml.replace(creditcardRegex, creditcardRedacted);

    console.info("[v04] Sanitized SUBJECT --->");
    console.info("[v04] " + sanitizedSubjectHtml);
    console.info("[v04] <--- Sanitized SUBJECT");

    console.info("[v04] Sanitized BODY --->");
    console.info("[v04] " + sanitizedBodyHtml);
    console.info("[v04] <--- Sanitized BODY");

    let promiseSetSubject = makePromiseSetSubject(item, sanitizedSubjectHtml);
    let promiseSetBody = makePromiseSetBody(item, sanitizedBodyHtml);

    Promise.all([promiseSetSubject, promiseSetBody]).then(() => {
      console.info("[v04] successfully set redacted SUBJECT / BODY:");

      // POST to LOG Server
      let logMessage = {
        mailbox: mailboxEmail,
        from: from,
        to: to,
        subject: subjectHtml,
        no_of_card: 5,
        no_of_nric: 10,
        random: new Date().toISOString()
      };

      fetch('https://demo-api.consentrade.io/api/v1/income', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': '9f4e1c3b7a6dbe2a4d85a9e7f1c23d9096a8b1f3c4d7e2a9c0b4d8f2e5a1c6b9'
        },
        body: JSON.stringify(logMessage),
      })
        .then(response => {
          console.info(`[v04] POST to LOG - fetch() response, STATUS=[${response.statusText}]`);

          if (!response.ok) {
            console.error(`[v04] POST to LOG - API failed, STATUS=[${response.statusText}]`);
            event.completed({ allowEvent: false, errorMessage: "Send LOG failed on API" });
          }

          console.info(`[v04] POST to LOG - API OK, STATUS=[${response.statusText}]`);

          return response.json();
        })
        .then(data => {
          console.info(`[v04] POST to LOG - API OK, DATA=[${data}]`);
          console.info(`[v04] POST to LOG - API OK, DATA JSON=[${JSON.stringify(data)}]`);

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
          console.error("[v04] POST to LOG - NETWORK failed :", error);
          event.completed({ allowEvent: false, errorMessage: "Send LOG failed on Network" });
        });

      // POST to LOG Server




      
    }).catch((error) => {
      console.error("[v04] An error occurred while setting item data:", error);
      event.completed({ allowEvent: false, errorMessage: "Not able to set TO, FROM, SUBJECT, BODY!!!" });
    });
  }).catch((error) => {
    console.error("[v04] An error occurred while fetching item data:", error);
    event.completed({ allowEvent: false, errorMessage: "Not able to retrieve TO, FROM, SUBJECT, BODY!!!" });
  });






  console.info("[v04] Commands.js::onMessageSendHandler(100)] Exit!");
}

