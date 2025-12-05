const VERSION = 'v005.003';


// ================= NRIC & CARDNO related =============================
const regexCreditCard = /\b(?:\d[ -]*?){13,16}\b/g;
const regexNRIC = /\b([SFTGM])(\d{7})([A-Z])\b/gi;

const nricRedacted = 'X0000000X';
const creditcardRedacted = 'xxxx-xxxx-xxxx-xxxx';

function isNricChecksumValid(match, prefix, digits, checksum) {
  //console.log(`[calculateChecksum] - match=[${match}], prefix=[${prefix}], digits=[${digits}], checksum=[${checksum}]`);

  const weights = [2, 7, 6, 5, 4, 3, 2];
  const nricChecksum = ['J', 'Z', 'I', 'H', 'G', 'F', 'E', 'D', 'C', 'B', 'A'];
  const finChecksum = ['X', 'W', 'U', 'T', 'R', 'Q', 'P', 'N', 'M', 'L', 'K'];

  prefix = prefix.toUpperCase();
  checksum = checksum.toUpperCase();

  let sum = 0;
  for (let i=0; i < digits.length; i++) {
    sum += digits[i] * weights[i];
  }
  //console.log(`[calculateChecksum] - sum=[${sum}]`);

  if (prefix === 'T' || prefix === 'G') {
    sum += 4;
  } else if (prefix === 'M') {
    sum += 3;
  }
  //console.log(`[calculateChecksum] - sum(adjusted)=[${sum}]`);

  const remainder = sum % 11;
  //console.log(`[calculateChecksum] - remainder=[${remainder}]`);


  let isValid = false;
  if (prefix === 'S' || prefix === 'T') {
    //console.log(`[calculateChecksum] - [S | T] should be [${nricChecksum[remainder]}] and pass-in is [${checksum}]`);
    if (checksum === nricChecksum[remainder]) {
        isValid = true;
    }
  } else if (prefix === 'F' || prefix === 'G' || prefix === 'M') {
    //console.log(`[calculateChecksum] - [F | G | M] should be [${finChecksum[remainder]}] and pass-in is [${checksum}]`);
    if (checksum === finChecksum[remainder]) {
        isValid = true;
    }
  }

  //console.log(`[calculateChecksum] - validation result [${isValid}]`);

  return isValid;
}

function countViolation(text) {
    let no_of_cardno = 0;
    const cardnoMatches = text.matchAll(regexCreditCard);
    const cardnoMatchesArray = Array.from(cardnoMatches);
    no_of_cardno = cardnoMatchesArray.length;
    // console.log(`CARDNO=[${no_of_cardno}]\n\n`);

    let no_of_nric = 0;
    const nricMatches = text.matchAll(regexNRIC);
    const nricMatchesArray = Array.from(nricMatches);
    //console.log(`Suspected NRIC=[${nricMatchesArray.length}]\n\n`);
    nricMatchesArray.forEach(nricMatch => {
        //console.log(`found [${nricMatch}]`);
        if (isNricChecksumValid(nricMatch[0], nricMatch[1], nricMatch[2], nricMatch[3])) {
            no_of_nric++;
        }
    });
    // console.log(`Actual NRIC=[${no_of_nric}]\n\n`);

    return no_of_cardno + no_of_nric;
}
// ================= NRIC & CARDNO related =============================


// Factories
const makePromiseSetSubject = (mailItem, newSubject) => {
  return new Promise((resolve, reject) => {
    console.log(`[makePromiseSetSubject()] setting SUBJECT to [${newSubject}]`);
    mailItem.subject.setAsync(newSubject, { coercionType: Office.CoercionType.subjectHtml }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`[makePromiseSetSubject()] set SUBJECT OK`);
        resolve(asyncResult.value);
      } else {
        console.log(`[makePromiseSetSubject()] set SUBJECT BAD : [${asyncResult.error.message}]`);
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseSetBody = (mailItem, newBody) => {
  return new Promise((resolve, reject) => {
    console.log(`[makePromiseSetSubject()] setting BODY to [${newBody}]`);
    mailItem.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`[makePromiseSetSubject()] set BODY OK`);
        resolve(asyncResult.value);
      } else {
        console.log(`[makePromiseSetSubject()] set BODY BAD : [${asyncResult.error.message}]`);
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetSubject = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.log(`[ARG] getting SUBJECT`);
    mailItem.subject.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.log(`[ARG] get SUBJECT OK`);
        resolve(asyncResult.value);
      } else {
        console.log(`[ARG] get SUBJECT BAD : [${asyncResult.error.message}]`);
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetBody = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.log(`[ARG] getting BODY`);
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

const makePromiseGetFrom = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info(`[ARG] getting FROM`);
    mailItem.from.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info(`[ARG] get FROM OK`);
        resolve(asyncResult.value);
      } else {
        console.info(`[ARG] get FROM BAD : [${asyncResult.error.message}]`);
        reject(asyncResult.error.message);
      }
    });
  });
};

const makePromiseGetTo = (mailItem) => {
  return new Promise((resolve, reject) => {
    console.info(`[ARG] getting TO`);
    mailItem.to.getAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        console.info(`[ARG] get TO OK`);

        const tos = asyncResult.value;
        let toList = "";
        for (let i=0; i<tos.length; i++) {
          if (i>0) {
            toList = toList + ", ";
          }
          toList = toList + tos[i].emailAddress;
        }
        resolve(toList);
      } else {
        console.info(`[ARG] get TO BAD : [${asyncResult.error.message}]`);
        reject(asyncResult.error.message);
      }
    });
  });
};



// Office.onReady();
Office.onReady((info) => {
  console.info(`[Office.onReady(${VERSION})]`);
  console.info(`[Office.onReady(${Office.context.platform})]`);
});

/**
 * The words in the subject or body that require corresponding color categories to be applied to a new
 * message or appointment.
 * @constant
 * @type {string[]}
 */
 const KEYWORDS = [
  "sales",
//  "expense reports",
//  "legal",
//  "marketing",
//  "performance reviews",
];

/**
 * Handle the OnNewMessageCompose or OnNewAppointmentOrganizer event by verifying that keywords have corresponding
 * color categories when a new message or appointment is created. If no corresponding categories exist, they will be
 * created.
 * @param {Office.AddinCommands.Event} event The OnNewMessageCompose or OnNewAppointmentOrganizer event object.
 */
function onItemComposeHandler(event) {
  console.info(`[commands.js::onItemComposeHandler()]`);




  /*
  Office.context.mailbox.masterCategories.getAsync(
    { asyncContext: event },
    (asyncResult) => {
      let event = asyncResult.asyncContext;

      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        event.completed({
          allowEvent: false,
          errorMessage: "Failed to configure categories.",
        });
        return;
      }

      let categories = asyncResult.value;
      let categoriesToBeCreated = [];
      if (categories) {
        let categoryNamesInUse = getCategoryProperty(categories, "displayName");
        let categoryColorsInUse = getCategoryProperty(categories, "color");
        categoriesToBeCreated = getCategoriesToBeCreated(
          KEYWORDS,
          categoryNamesInUse
        );

        if (categoriesToBeCreated.length > 0) {
          categoriesToBeCreated = assignCategoryColors(
            categoriesToBeCreated,
            categoryColorsInUse
          );
        }
      } else {
        categoriesToBeCreated = assignCategoryColors(
          getCategoriesToBeCreated(KEYWORDS)
        );
      }

      createCategories(event, categoriesToBeCreated);
      event.completed({ allowEvent: true });
    }
  );
  */
}



function action(event) {
  console.info(`[action(start)]`);
  //console.info(`[action()] - EVENT=[${JSON.stringify(event)}]`);

  

  if (event.context) {
    const context = JSON.parse(event.context);
    const caller = context.caller;
    if (caller) {
      console.info(`[action()] - CALLER=[${caller}]`);

      if (caller === 'onItemSendHandler()') {
        console.info(`[action()] - caller [${caller}] --- !!!ACTION!!!`);

        const item = Office.context.mailbox.item;


        item.loadCustomPropertiesAsync((asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to load1 custom properties: " + asyncResult.error.message);
            return;
          }

          const customProps = asyncResult.value;

          // Set a new property on the item.
          customProps.set("UserPermission", "--==>Denied<==--");

          // Save the properties back to the server.
          customProps.saveAsync((saveResult) => {
            if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Custom property saved successfully.");

              item.loadCustomPropertiesAsync((asyncResult2) => {
                if (asyncResult2.status === Office.AsyncResultStatus.Succeeded) {
                  console.info(`Succeeded to load2 custom properties: [${JSON.stringify(asyncResult2.value)}]`);
                } else {
                  console.error("Failed to load2 custom properties: " + asyncResult2.error.message);
                  return;
                }



                item.subject.getAsync((asyncResult) => {
                  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log(`[action()] SUBJECT=[${asyncResult.value}]`);
                    const newSubject = `[REDACTED]-[${Office.context.platform}]: ${asyncResult.value}`;


                    item.subject.setAsync(newSubject, { coercionType: Office.CoercionType.subjectHtml }, function (asyncResult) {
                      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log(`[makePromiseSetSubject()] set SUBJECT OK`);

                        item.sendAsync(function (asyncResult) {
                          /*
                          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            console.info(`[action()] : sendAsync() Succeeded`);
                          } else {
                            console.info(`[action()] : sendAsync() Failed`);
                          }
                          event.completed({ allowEvent: true, context: JSON.stringify({ a: 'called by action()' }), });
                          return;
                          */
                        });


                      } else {
                        console.log(`[makePromiseSetSubject()] Set SUBJECT BAD : [${asyncResult.error.message}]`);
                      }
                    });
                  } else {
                    console.log(`[action()] Get SUBJECT BAD : [${asyncResult.error.message}]`);
                  }
                });
              });



            } else {
              console.error("Failed to save custom properties: " + saveResult.error.message);
              
            }
          });
        });

        




      } else {
        console.info(`[action()] - unidentified caller [${caller}]`);
      }
    } else {
      console.info(`[action(4002)] - no caller`);
    }
  } else {
    console.info(`[action(4001)] - no context`);
  }
}


/**
 * Handle the OnMessageSend or OnAppointmentSend event by verifying that applicable color categories are
 * applied to a new message or appointment before it's sent.
 * @param {Office.AddinCommands.Event} event The OnMessageSend or OnAppointmentSend event object.
 */
/**
    1. get subject
    2. extract keywords from subject
    3. fetch email body
    4. check applied categories - checkAppliedCategories(event, detectedWords);
 */
function onItemSendHandler(event) {
  console.info(`[onItemSendHandler()]`);
  console.info(`[onItemSendHandler()] - EVENT=[${JSON.stringify(event)}]`);

  const item = Office.context.mailbox.item;
  item.subject.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const subject = asyncResult.value + "";
      console.log(`[onItemSendHandler()] SUBJECT=[${subject}]`);


      item.loadCustomPropertiesAsync((asyncLoadResult) => {
        if (asyncLoadResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to load custom properties: " + asyncLoadResult.error.message);
          return;
        }

        const customProps = asyncLoadResult.value;

        const myData = customProps.get("UserPermission");
        // customProps.set("UserPermission", "Granted");

        if (myData) {
          console.log("Retrieved custom data: " + myData);
        } else {
          console.log("Custom data not found.");
        }

        if (subject.startsWith("[REDACTED]")) {
          console.log(`[onItemSendHandler()] SUBJECT REDACTED, SEND!!!`);


          const details = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            icon: "Icon.32x32",
            message: `Redacted all NRIC/Credit card Number`,
            persistent: true
          };
          Office.context.mailbox.item.notificationMessages.addAsync('redacter', details, (asyncNotificationResult) => {
            if (asyncNotificationResult.status === Office.AsyncResultStatus.Succeeded) {
              console.info("asyncNotificationResult OK");
            } else {
              console.info("asyncNotificationResult ERROR");
            }

            event.completed({
              allowEvent: true,
            });
            return;
          });
          
          
        } else {
          console.log(`[onItemSendHandler()] SUBJECT NOT REDACTED, go react!!!`);
          event.completed({
            allowEvent: false,
            errorMessage: `[onItemSendHandler(1000): [${Office.context.platform}]]`,
            errorMessageMarkdown: '[onItemSendHandler(Markdown:1000)]\n\nHold-on.\n\n**Tip**: we are [${Office.context.platform}]...',
            commandId: "ActionButton",
            contextData: JSON.stringify({ caller: 'onItemSendHandler()' }),
          });
        }
      });


      






    } else {
      console.log(`[onItemSendHandler()] Get SUBJECT BAD : [${asyncResult.error.message}]`);
      event.completed({
          allowEvent: false,
          errorMessage: "[onItemSendHandler()] Failed to fetch SUBJECT",
      });
      return;
    }
  });


  


  
  // cancelLabel: "Redact & Send",
}

/**
 * Get the property values of existing categories.
 * @param {Office.CategoryDetails[]} categories Existing categories in Outlook.
 * @param {string} property The property to extract from existing categories. Categories have a display name and a color.
 * @returns {string[]} The property's value.
 */
function getCategoryProperty(categories, property) {
  console.info(`[commands.js::getCategoryProperty()]`);


  let values = [];
  categories.forEach((category) => {
    values.push(category[property]);
  });

  console.info(`[commands.js::getCategoryProperty()] - BODY=[${values}]`);

  return values;
}

/**
 * Determine the categories to be created based on existing categories.
 * @param {string[]} keywords The keywords that require corresponding categories.
 * @param {string[]} existingCategories The display names currently in use by existing categories.
 * @returns {string[]} The names of the new categories.
 */
function getCategoriesToBeCreated(keywords, existingCategories = []) {
  console.info(`[commands.js::getCategoriesToBeCreated()]`);


  let categoriesToBeCreated = [];
  if (existingCategories.length === 0) {
    keywords.forEach((word) => {
      categoriesToBeCreated.push(`Office Add-ins Sample: ${word}`);
    });
  } else {
    keywords.forEach((word) => {
      if (!existingCategories.includes(`Office Add-ins Sample: ${word}`)) {
        categoriesToBeCreated.push(`Office Add-ins Sample: ${word}`);
      }
    });
  }

  console.info(`[commands.js::getCategoriesToBeCreated()] - BODY=[${categoriesToBeCreated}]`);

  return categoriesToBeCreated;
}

/**
 * Assign a color to a new category based on available colors. If all 25 colors are in use,
 * duplicate colors are assigned starting from Preset0.
 * @param {string[]} categoriesToBeCreated The names of the new categories.
 * @param {string[]} categoryColorsInUse The colors currently in use by existing categories.
 * @returns {Office.CategoryDetails[]} The new category objects to be created.
 */
function assignCategoryColors(
  categoriesToBeCreated = [],
  categoryColorsInUse = []
) {
  console.info(`[commands.js::assignCategoryColors()]`);


  const totalColors = 25;
  if (categoryColorsInUse.length >= totalColors) {
    for (let i = 0; i < categoriesToBeCreated.length; i++) {
      categoriesToBeCreated[i] = {
        displayName: categoriesToBeCreated[i],
        color: `Preset${i}`,
      };
    }
  } else {
    for (let i = 0; i < categoriesToBeCreated.length; i++) {
      for (let j = 0; j < totalColors; j++) {
        if (!categoryColorsInUse.includes(`Preset${j}`)) {
          categoriesToBeCreated[i] = {
            displayName: categoriesToBeCreated[i],
            color: `Preset${j}`,
          };

          categoryColorsInUse.push(`Preset${j}`);
          break;
        }
      }
    }
  }

  console.info(`[commands.js::assignCategoryColors()] - BODY=[${categoriesToBeCreated}]`);

  return categoriesToBeCreated;
}

/**
 * Create categories.
 * @param {Office.AddinCommands.Event} event The OnNewMessageCompose or OnNewAppointmentOrganizer event object.
 * @param {Office.CategoryDetails[]} categoriesToBeCreated The new category objects to create.
 */
function createCategories(event, categoriesToBeCreated) {
  console.info(`[commands.js::createCategories()]`);


  Office.context.mailbox.masterCategories.addAsync(
    categoriesToBeCreated,
    { asyncContext: event },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        asyncResult.asyncContext.completed({
          allowEvent: false,
          errorMessage: "Failed to set new categories.",
        });
        return;
      }
    }
  );
}

/**
 * Determine if keywords are present in the message or appointment's subject or body that require corresponding categories.
 * @param {string[]} keywords The keywords that require corresponding categories.
 * @param {string} text The contents of the subject or body of the message or appointment.
 * @param {string[]} detectedWords The keywords found in the message or appointment's subject or body.
 * @returns {string[]} Keywords detected in the message or appointment's subject or body that require corresponding categories.
 */
function checkForKeywords(keywords, text, detectedWords = []) {
  keywords = new RegExp(keywords.join("|"), "gi");
  text = text.toLowerCase();

  let keywordsFound = text.match(keywords);
  if (keywordsFound) {
    checkForDuplicates(keywordsFound, detectedWords);
  }

  return detectedWords;
}

/**
 * Check for duplicate keywords in the message or appointment's subject or body.
 * @param {string[]} wordsToCompare The keywords found in the message or appointment's subject or body to compare to the existing
 * list of detected keywords.
 * @param {string[]} wordList The existing list of detected keywords.
 */
function checkForDuplicates(wordsToCompare = [], wordList = []) {
  wordsToCompare.forEach((word) => {
    if (!wordList.includes(word)) {
      wordList.push(word);
    }
  });
}

/**
 * Determine the categories to be added based on the detected keywords in the message or appointment's subject or body.
 * @param {string[]} detectedWords The keywords detected in the message or appointment's subject or body.
 * @returns {string[]} The names of the categories to be added to the message or appointment.
 */
function getCategoryName(detectedWords) {
  let categories = [];
  detectedWords.forEach((word) => {
    categories.push(`Office Add-ins Sample: ${word}`);
  });

  return categories;
}

/**
 * Check that the appropriate categories, based on detected keywords in the subject or body, are applied to the
 * message or appointment before it's sent.
 * @param {Office.AddinCommands.Event} event The OnMessageSend or OnAppointmentSend event object.
 * @param {string[]} detectedWords The keywords found in the message or appointment's subject or body.
 */
function checkAppliedCategories(event, detectedWords) {
  console.info(`[commands.js::checkAppliedCategories()]`);


  let options = {
    asyncContext: { callingEvent: event, keywordArray: detectedWords },
  };
  Office.context.mailbox.item.categories.getAsync(options, (asyncResult) => {
    let sendEvent = asyncResult.asyncContext.callingEvent;

    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error(asyncResult.error.message);
      sendEvent.completed({
        allowEvent: false,
        errorMessage: "[commands.js::checkAppliedCategories(2001)] - Failed to check categories applied to the item.",
      });
      return;
    }

    let requiredCategories = getCategoryName(
      asyncResult.asyncContext.keywordArray
    );
    let detectedCategories = asyncResult.value;
    if (detectedCategories) {
      let detectedCategoryNames = getCategoryProperty(
        detectedCategories,
        "displayName"
      );
      let missingCategories = getMissingCategories(
        requiredCategories,
        detectedCategoryNames
      );
      if (missingCategories.length > 0) {
        let message = `Don't forget to also add the following categories: ${missingCategories.join(", ")}`;
        console.log(message);
        sendEvent.completed({ allowEvent: false, errorMessage: message });
        return;
      }

      sendEvent.completed({ allowEvent: true });
    } else {
      let message = `You must assign the following categories before your ${
        Office.context.mailbox.item.itemType
      } can be sent: ${requiredCategories.join(", ")}`;
      console.log(message);
      sendEvent.completed({ allowEvent: false, errorMessage: message });
      return;
    }
  });
}

/**
 * Get the names of the required categories still missing from the message or appointment.
 * @param {string[]} requiredCategories The names of the categories required on the message or appointment before it can be sent.
 * @param {string[]} appliedCategories The names of the categories that are currently applied to the message or appointment.
 * @returns {string[]} The names of the categories that need to be applied to the message or appointment.
 */
function getMissingCategories(requiredCategories, appliedCategories) {
  let missingCategories = requiredCategories.filter(
    (category) => !appliedCategories.includes(category)
  );
  return missingCategories;
}

Office.actions.associate("onMessageComposeHandler", onItemComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onItemComposeHandler);
Office.actions.associate("onMessageSendHandler", onItemSendHandler);
Office.actions.associate("onAppointmentSendHandler", onItemSendHandler);
Office.actions.associate("ActionButton", action);