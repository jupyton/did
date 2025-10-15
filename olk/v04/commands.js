const MY_NAME = 'v04 - 002';


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
    const mailboxEmail = userProfile.emailAddress;
    if (mailboxEmail) {
      console.info(`[v04] SENDER MAILBOX=[${mailboxEmail}]`);
    } else {
      console.err("[v04] SENDER MAILBOX not available.");
    }
  } else {
    console.err("[v04] SENDER MAILBOX not available.");
  }

  //Office.context.mailbox.item.to.getAsync(function (asyncResult) {
  //  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //    const msgTo = asyncResult.value;
  //    console.info("[v04] 2. Message being sent to:");
  //    for (let i = 0; i < msgTo.length; i++) {
  //      console.info(msgTo[i].displayName + " (" + msgTo[i].emailAddress + ")");
  //    }
  //  } else {
  //    console.error("[v04] 2. ERROR while trying to get TO field");
  //    console.error(asyncResult.error);
  //  }
  //});

  //console.info("[v04] 3. trying to get FROM field");
  //Office.context.mailbox.item.from.getAsync(function (asyncResult) {
  //  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
  //    const msgFrom = asyncResult.value;
  //    console.info("[v04] 3. from-from-from Message FROM: " + msgFrom.displayName + " (" + msgFrom.emailAddress + ")");
  //  } else {
  //    console.error("[v04] 3. from-from-from ERROR while trying to get FROM field");
  //    console.error(asyncResult.error);
  //  }
  //});

  // ======== get Identity



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

