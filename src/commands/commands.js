/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Ensures the Office.js library is loaded.
 */
Office.onReady((info) => {
  /** 
   * Maps the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
   * This ensures support in Outlook on Windows. 
   */
  if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }
});

/**
 * The legal hold email account of the fictitious company, Fabrikam. It's added to the Bcc field of a
 * message that's configured with the Highly Confidential sensitivity label.
 * @constant
 * @type {string}
 */
const LEGAL_HOLD_ACCOUNT = "gpec@marr.it";

/**
 * The email address suffix that identifies an account owned by a legal team member at Fabrikam.
 * @constant
 * @type {string}
 */
const PEC_SUFFIX = "gpec*@marr.it";


/**
 * Handle the OnMessageSend event by checking whether the current message has an attachment or a recipient is a member
 * of the legal team. If either of these conditions is true, the event handler checks for the Highly Confidential sensitivity
 * label on the message and sets it if needed.
 * @param {Office.AddinCommands.Event} event The OnMessageSend event object. 
 */
function onMessageSendHandler(event) {
  
  console.log("Version 02");

  Office.context.mailbox.item.from.getAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to get the Sender from the From field.");
      console.log(`Error: ${result.error.message}`);
      event.completed({ allowEvent: false, errorMessage: "Unable to get the recipients from the From field. Save your message, then restart Outlook." });
      return;
    }

    if (containsPecSender(result.value)) {
        event.completed({ allowEvent: false, errorMessage:  "OK" });
    } else {
      event.completed({ allowEvent: false, errorMessage:  "KO" });
    }
  });
}


/**
 * Check that the Highly Confidential sensitivity label is set if a message contains an attachment or a recipient
 * who's a member of the legal team.
 * @param {Office.AddinCommands.Event} event The OnMessageSend event object.
 */
function ensureHighlyConfidentialLabelSet(event) {
  Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve the status of the sensitivity label catalog.");
      console.log(`Error: ${result.error.message}`);
      event.completed({ allowEvent: false, errorMessage: "Unable to retrieve the status of the sensitivity label catalog. Save your message, then restart Outlook." });
      return;
    }

    Office.context.sensitivityLabelsCatalog.getAsync({ asyncContext: event }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to retrieve the catalog of sensitivity labels.");
        console.log(`Error: ${result.error.message}`);
        event.completed({ allowEvent: false, errorMessage: "Unable to retrieve the catalog of sensitivity labels. Save your message, then restart Outlook." });
        return;
      }

      const highlyConfidentialLabel = getLabelId("Highly Confidential", result.value);
      Office.context.mailbox.item.sensitivityLabel.getAsync({ asyncContext: { event: event, highlyConfidentialLabel: highlyConfidentialLabel } }, (result) => {
        const event = result.asyncContext.event;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to get the sensitivity label of the message.");
          console.log(`Error: ${result.error.message}`);
          event.completed({ allowEvent: false, errorMessage: "Unable to get the sensitivity label applied to the message. Save your message, then restart Outlook." });
          return;
        }

        const highlyConfidentialLabel = result.asyncContext.highlyConfidentialLabel;
        if (result.value === highlyConfidentialLabel) {
          event.completed({ allowEvent: true });
        } else {
          Office.context.mailbox.item.sensitivityLabel.setAsync(highlyConfidentialLabel, { asyncContext: event }, (result) => {
            const event = result.asyncContext;
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.log("Unable to set the Highly Confidential sensitivity label to the message.");
              console.log(`Error: ${result.error.message}`);
              event.completed({ allowEvent: false, errorMessage: "Unable to set the Highly Confidential sensitivity label to the message. Save your message, then restart Outlook." });
              return;
            }

            event.completed({ allowEvent: false, errorMessage: "Due to the contents of your message, the sensitivity label has been set to Highly Confidential and the Legal Hold account has been added to the Bcc field.\nTo learn more, see Fabrikam's information protection policy.\n\nDo you need to make changes to your message?" });
          });
        }
      });
    });
  });
}

/**
 * Check whether the legal hold account was added to the Bcc field if the sensitivity label of a message is set to
 * Highly Confidential. If the account appears in the Bcc field, but the sensitivity label isn't set to
 * Highly Confidential, the account is removed from the message.
 * @param {Office.AddinCommands.Event} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
 */
function checkForLegalHoldAccount(event) {
  Office.context.sensitivityLabelsCatalog.getIsEnabledAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve the status of the sensitivity label catalog.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    Office.context.sensitivityLabelsCatalog.getAsync({ asyncContext: event }, (result) => {
      const event = result.asyncContext;
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Unable to retrieve the catalog of sensitivity labels.");
        console.log(`Error: ${result.error.message}`);
        event.completed();
        return;
      }

      const highlyConfidentialLabel = getLabelId("Highly Confidential", result.value);
      Office.context.mailbox.item.sensitivityLabel.getAsync({ asyncContext: { event: event, highlyConfidentialLabel: highlyConfidentialLabel, } }, (result) => {
        const event = result.asyncContext.event;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to get the sensitivity label of the message.");
          console.log(`Error: ${result.error.message}`);
          event.completed();
          return;
        }

        if (result.value === result.asyncContext.highlyConfidentialLabel) {
          addLegalHoldAccount(event, Office.context.mailbox.item.bcc);
        } else {
          removeLegalHoldAccount(event, Office.context.mailbox.item.bcc);
        }
      });
    });
  });
}

/**
 * Get the index of the legal hold account in the To, Cc, or Bcc field.
 * @param {Office.EmailAddressDetails[]} recipients The recipients in the To, Cc, or Bcc field.
 * @returns {number} The index of the legal hold account.
 */
function getLegalHoldAccountIndex(recipients) {
  return recipients.findIndex((recipient) => (recipient.emailAddress).toLowerCase() === LEGAL_HOLD_ACCOUNT);
}

/**
 * Remove the legal hold email account from the To, Cc, or Bcc field of a message.
 * @param {Office.AddinCommands.Event} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
 * @param {Office.Recipients} recipientField The recipient object of the To, Cc, or Bcc field of a message.
 */
function removeLegalHoldAccount(event, recipientField) {
  recipientField.getAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve recipients from the field.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    const recipients = result.value;
    const legalHoldAccountIndex = getLegalHoldAccountIndex(recipients);
    if (legalHoldAccountIndex > -1) {
      recipients.splice(legalHoldAccountIndex, 1);
      recipientField.setAsync(recipients, { asyncContext: event }, (result) => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log("Unable to set the recipients.");
          console.log(`Error: ${result.error.message}`);
          event.completed();
          return;
        }
    
        console.log(`${LEGAL_HOLD_ACCOUNT} has been removed.`);
        event.completed();
      });
    }
  });
}

/**
 * Add the legal hold email account to the Bcc field.
 * @param {Office.AddinCommands.Event} event The OnMessageRecipientsChanged or OnSensitivityLabelChanged event object.
 * @param {Office.Recipients} recipientField The recipient object of the Bcc field.
 */
function addLegalHoldAccount(event, recipientField) {
  recipientField.getAsync({ asyncContext: event }, (result) => {
    const event = result.asyncContext;
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.log("Unable to retrieve recipients from the field.");
      console.log(`Error: ${result.error.message}`);
      event.completed();
      return;
    }

    const recipients = result.value;
    const legalHoldAccountIndex = getLegalHoldAccountIndex(recipients);
    if (legalHoldAccountIndex === -1) {
      recipientField.addAsync([LEGAL_HOLD_ACCOUNT], { asyncContext: event }, (result) => {
        const event = result.asyncContext;
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.log(`Unable to add ${LEGAL_HOLD_ACCOUNT} as a recipient.`);
          console.log(`Error: ${result.error.message}`);
          event.completed();
          return;
        }

        console.log(`${LEGAL_HOLD_ACCOUNT} has been added to the Bcc field.`);
        event.completed();
      });
    }
  });
}

/**
 * Get the unique identifier (GUID) of a sensitivity label.
 * @param {string} sensitivityLabel The name of a sensitivity label.
 * @param {Office.SensitivityLabelDetails[]} sensitivityLabelCatalog The catalog of sensitivity labels.
 * @returns {number} The GUID of a sensitivity label. 
 */
function getLabelId(sensitivityLabel, sensitivityLabelCatalog) {
  return (sensitivityLabelCatalog.find((label) => label.name === sensitivityLabel)).id;
}

/**
 * Check if a member of the PEC team is a recipient in the To, Cc, or Bcc field.
 * @param {Office.EmailAddressDetails[]} recipients The recipients in the To, Cc, or Bcc field.
 * @returns {boolean} Returns true if a member of the legal team is a recipient.
 */
function containsPecSender(recipients) {
  console.log("recipients");
  console.log(recipients );
  console.log( recipients.length);
  for (let i = 0; i < recipients.length; i++) {
    console.log(`Error: ${recipients[i].emailAddress.toLowerCase()}`);

    const emailAddress = recipients[i].emailAddress.toLowerCase();
    if (emailAddress.includes(PEC_SUFFIX)) {
      return true;
    }
  }

  return false;
}