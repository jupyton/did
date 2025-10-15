const MY_NAME = 'v05 - 007';

const ALLOW_ENTRIES = 5;
const regexCreditCard = /\b(?:\d[ -]*?){13,16}\b/g;
const regexNRIC = /\b([SFTGM])(\d{7})([A-Z])\b/gi;

const ssnRegex = /\b(\d{3}-\d{2}-\d{4}|\d{9})\b/g;
const nricRegex = /[SFTG]\d{7}[A-Z]/gm;
const creditcardRegex = /(\d{4}[-]){3}\d{4}|\d{16}/gm;



const nricRedacted = 'X0000000X';
const creditcardRedacted = 'xxxx-xxxx-xxxx-xxxx';


Office.onReady((info) => {
  console.info(`[ARG] ${MY_NAME} Commands.js::onReady()`);
});

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);

// validate NRIC checksum
function validateChecksum(match, prefix, digits, checksum) {
  const weights = [2, 7, 6, 5, 4, 3, 2];
  const nricChecksum = ['J', 'Z', 'I', 'H', 'G', 'F', 'E', 'D', 'C', 'B', 'A'];
  const finChecksum = ['X', 'W', 'U', 'T', 'R', 'Q', 'P', 'N', 'M', 'L', 'K'];

  prefix = prefix.toUpperCase();
  checksum = checksum.toUpperCase();

  let sum = 0;
  for (let i=0; i < digits.length; i++) {
    sum += digits[i] * weights[i];
  }

  if (prefix === 'T' || prefix === 'G') {
    sum += 4;
  } else if (prefix === 'M') {
    sum += 3;
  }

  const remainder = sum % 11;

  let isValid = false;
  if (prefix === 'S' || prefix === 'T') {
    if (checksum === nricChecksum[remainder]) {
        isValid = true;
    }
  } else if (prefix === 'F' || prefix === 'G' || prefix === 'M') {
    if (checksum === finChecksum[remainder]) {
        isValid = true;
    }
  }

  console.log(`[validateChecksum] - pass-in [${prefix}, ${digits}, ${checksum}], calculated=[${finChecksum[remainder]}], result=[${isValid}]`);

  return isValid;
}


// Factories
const makePromiseSetSubject = (mailItem, newSubject) => {
  return new Promise((resolve, reject) => {
    console.info("[ARG] create PROMISE to SET SUBJECT to [" + newSubject + "]");
    mailItem.subject.setAsync(newSubject, { coercionType: Office.CoercionType.subjectHtml }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[ARG] SET SUBJECT OK : asyncResult-Value=[" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.info("[ARG] SET SUBJECT BAD : " + setAsyncResult.error.message + "]");
        reject(setAsyncResult.error.message);
      }
    });
  });
};

const makePromiseSetBody = (mailItem, newBody) => {
  return new Promise((resolve, reject) => {
    console.info("[ARG] create PROMISE to SET BODY to [" + newBody + "]");
    mailItem.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, function (setAsyncResult) {
      if (setAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[ARG] SET BODY OK : asyncResult-Value=[" + setAsyncResult.value + "]");
        resolve(setAsyncResult.value);
      } else {
        console.info("[ARG] SET BODY BAD : [" + setAsyncResult.error.message + "]");
        reject(setAsyncResult.error.message);
      }
    });
  });
};


const makePromiseGetSubject = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info("[ARG] create PROMISE to GET SUBJECT");
    mailItem.subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[ARG] GET SUBJECT OK");
        resolve(asyncResult.value);
      } else {
        console.info("[ARG] GET SUBJECT BAD");
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetBody = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info("[ARG] create PROMISE to GET BODY");
    mailItem.body.getAsync(Office.CoercionType.Html, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[ARG] GET BODY OK");
        resolve(asyncResult.value);
      } else {
        console.info("[ARG] GET BODY BAD");
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetTo = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info("[ARG] create PROMISE to GET TO");
    mailItem.to.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[ARG] GET TO OK");
        resolve(asyncResult.value);
      } else {
        console.info("[ARG] GET TO BAD");
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetFrom = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info("[ARG] create PROMISE to GET FROM");
    mailItem.from.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info("[ARG] GET FROM OK");
        resolve(asyncResult.value);
      } else {
        console.info("[ARG] GET FROM BAD");
        reject(asyncResult.error.message);
      }
    });
  });
};


function onMessageSendHandler(event) {
  console.info("[ARG] Commands.js::onMessageSendHandler(): Received OnMessageSend event!");


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
      console.info(`[ARG] SENDER MAILBOX=[${mailboxEmail}]`);
    } else {
      console.err("[ARG] SENDER MAILBOX not available.");
    }
  } else {
    console.err("[ARG] SENDER MAILBOX not available.");
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

    console.info("[ARG] MAILBOX --->");
    console.info(`[ARG]  - [${mailboxEmail}]`);
    console.info("[ARG] <--- MAILBOX");

    console.info("[ARG] FROM --->");
    console.info(`[ARG]  - [${from.displayName} - ${from.emailAddress}]`);
    console.info("[ARG] <--- FROM");

    console.info("[ARG] TO --->");
    for (let i = 0; i < to.length; i++) {
      console.info(`[ARG]  - [${to[i].displayName} - ${to[i].emailAddress}]`);
    }
    console.info("[ARG] <--- TO");

    console.info("[ARG] SUBJECT --->");
    console.info("[ARG] " + subjectHtml);
    console.info("[ARG] <--- SUBJECT");

    console.info("[ARG] BODY --->");
    console.info("[ARG] " + bodyHtml);
    console.info("[ARG] <--- BODY");


    // Redacting NRIC
    sanitizedSubjectHtml = subjectHtml.replaceAll(regexNRIC, (match, prefix, digits, checksum) => {
      console.log(`checking - match=[${match}], prefix=[${prefix}], digits=[${digits}], checksum=[${checksum}]`);
      const isValid = validateChecksum(match, prefix, digits, checksum);

      if (isValid) {
        const redactedDigits = digits.slice(0, 4).replace(/\d/g, 'x') + digits.slice(4);
        return prefix + redactedDigits + checksum;
      }

      return match;
    });
    sanitizedBodyHtml = bodyHtml.replaceAll(regexNRIC, (match, prefix, digits, checksum) => {
      console.log(`checking - match=[${match}], prefix=[${prefix}], digits=[${digits}], checksum=[${checksum}]`);
      const isValid = validateChecksum(match, prefix, digits, checksum);

      if (isValid) {
        const redactedDigits = digits.slice(0, 4).replace(/\d/g, 'x') + digits.slice(4);
        return prefix + redactedDigits + checksum;
      }

      return match;
    });

    // Redacting Credit Card Numbers
    sanitizedSubjectHtml = sanitizedSubjectHtml.replaceAll(regexCreditCard, (match) => { 
      const digits = match.replace(/[- ]/g, '');
      const lastFour = digits.slice(-4);
      return `****-****-****-${lastFour}`;
    });
    sanitizedBodyHtml = sanitizedBodyHtml.replaceAll(regexCreditCard, (match) => { 
      const digits = match.replace(/[- ]/g, '');
      const lastFour = digits.slice(-4);
      return `****-****-****-${lastFour}`;
    });


    console.info("[ARG] Sanitized SUBJECT --->");
    console.info("[ARG] " + sanitizedSubjectHtml);
    console.info("[ARG] <--- Sanitized SUBJECT");

    console.info("[ARG] Sanitized BODY --->");
    console.info("[ARG] " + sanitizedBodyHtml);
    console.info("[ARG] <--- Sanitized BODY");

    let promiseSetSubject = makePromiseSetSubject(item, sanitizedSubjectHtml);
    let promiseSetBody = makePromiseSetBody(item, sanitizedBodyHtml);

    Promise.all([promiseSetSubject, promiseSetBody]).then(() => {
      console.info("[ARG] successfully set redacted SUBJECT / BODY:");

      const max = 1000000000000;
      const randomInteger = Math.round(Math.random() * max) + max;
      const ut = new Date().getTime();
      const nounce = randomInteger + ut;

      // POST to LOG Server
      let logMessage = {
        mailbox: mailboxEmail,
        from: from,
        to: to,
        subject: subjectHtml,
        no_of_card: 5,
        no_of_nric: 10,
        nounce: nounce
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
          console.info(`[ARG] POST to LOG - fetch() response, STATUS=[${response.statusText}]`);

          if (response.ok) {
            console.info(`[ARG] POST to LOG - API Response OK]`);
            //console.info(`[ARG] POST to LOG - API OK, DATA JSON=[${JSON.stringify(data)}]`);

            Office.context.mailbox.item.notificationMessages.replaceAsync('redacter', {
              type: 'errorMessage',
              message: `Everything OK, but still don't let you send [${MY_NAME}]`
            }, function (result) {
            });

            event.completed({
              allowEvent: false,
              errorMessage: "Everything OK, but still don't let you send"
            });

          } else {
            console.error(`[ARG] POST to LOG - API Response failed]`);
            event.completed({ allowEvent: false, errorMessage: "Send LOG failed on API" });

            Office.context.mailbox.item.notificationMessages.replaceAsync('redacter', {
              type: 'errorMessage',
              message: `Send LOG failed on API [${MY_NAME}]`
            }, function (result) {
            });

            event.completed({
              allowEvent: false,
              errorMessage: "Send LOG failed on API"
            });
          }
        })
        .catch(error => {
          console.error("[ARG] POST to LOG - NETWORK failed :", error);
          event.completed({ allowEvent: false, errorMessage: "Send LOG failed on Network" });
        });

      // POST to LOG Server




      
    }).catch((error) => {
      console.error("[ARG] An error occurred while setting item data:", error);
      event.completed({ allowEvent: false, errorMessage: "Not able to set TO, FROM, SUBJECT, BODY!!!" });
    });
  }).catch((error) => {
    console.error("[ARG] An error occurred while fetching item data:", error);
    event.completed({ allowEvent: false, errorMessage: "Not able to retrieve TO, FROM, SUBJECT, BODY!!!" });
  });






  console.info("[v04] Commands.js::onMessageSendHandler(100)] Exit!");
}

