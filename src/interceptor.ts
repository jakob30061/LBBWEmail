let item;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    item = Office.context.mailbox.item;
  }
});

function sendEmailInterceptor(event) {
  item.body.getAsync("text", { asyncContext: event }, checkEmail);
}

function checkEmail(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    const body = asyncResult.value;

    const strings = {
      attachments: 'Die Email scheint einen großen Anhang zu haben',
      insufficient_content: 'Die Email scheint über wenig Inhalt verfügen',
      many_recipient: 'Die Email hat viele Empfänger.'
    }

    const message = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: ''
    };

    if(message.message === '') {
      asyncResult.asyncContext.completed({ allowEvent: true });
      return
    }

    message.message = message.message + ' Ist diese Email wirklich relevant?'
    item.notificationMessages.replaceAsync("ActionPerformanceNotification", message);
    asyncResult.asyncContext.completed({ allowEvent: false });
  }

  
}

function emailHasManyRecipients() {
  let toRecipients, ccRecipients, bccRecipients;

  // Verify if the mail item is an appointment or message.
  if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
    toRecipients = item.requiredAttendees;
    ccRecipients = item.optionalAttendees;
  }
  else {
    toRecipients = item.to;
    ccRecipients = item.cc;
    bccRecipients = item.bcc;
  }

  // Get the recipients from the To or Required field of the item being composed.
  return toRecipients.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed)
      return false;
    
    return asyncResult.value.length < 30
  });
}