/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/* eslint-disable office-addins/no-office-initialize */

let mailboxItem: Office.MessageCompose;

Office.initialize = function (reason) {
    mailboxItem = Office.context.mailbox.item;
}

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

export const g = getGlobal() as any;

// The ui-less functions need to be available in global scope
g.validateEmailAddresses = validateEmailAddresses;

let addressesDialog: Office.Dialog;
function validateEmailAddresses(event) {
    let url: string = `${window.location.origin}/dialog.html`;
    Office.context.ui.displayDialogAsync(url, getDialogOptions(), dialogResult => {
        addressesDialog = dialogResult.value;
        addressesDialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => receiveMessage(message, event));
        addressesDialog.addEventHandler(Office.EventType.DialogEventReceived, () => dialogClosed(event));
    });

/*     mailboxItem.body.getAsync("text", { asyncContext: event }, result => {

        
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            showNotification = result.value.length < 30

            const strings = {
              attachments: 'Die Email scheint einen großen Anhang zu haben',
              insufficient_content: 'Die Email scheint über wenig Inhalt verfügen',
              many_recipient: 'Die Email hat viele Empfänger.'
            }

            return mailboxItem.to.getAsync((asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed)
                  return showNotification = showNotification ? showNotification : false;
                
                showNotification = showNotification ? showNotification : asyncResult.value.length < 30

                if(showNotification) {
                    let url: string = `${window.location.origin}/dialog.html`;
    
                    Office.context.ui.displayDialogAsync(url, getDialogOptions(), dialogResult => {
                        addressesDialog = dialogResult.value;
                        addressesDialog.addEventHandler(Office.EventType.DialogMessageReceived, (message) => receiveMessage(message, event));
                        addressesDialog.addEventHandler(Office.EventType.DialogEventReceived, () => dialogClosed(event));
                    });
                }
            });
        }
    }); */
}

function getDialogOptions(): Office.DialogOptions {
    let dialogOptions: Office.DialogOptions;
    // height: height of the dialog as a percentage of the current display. Defaults to 80%. 250px minimum
    // width: width of the dialog as a percentage of the current display. Defaults to 80%. 150px minimum

    if (Office.context.diagnostics.platform === Office.PlatformType.OfficeOnline) { // Browser

        // On Browser the iframe cannot get the dimensions of the parent window since it is on another domain
        dialogOptions = { width: 35, height: 9, displayInIframe: true }; // Browser

    } else { // Desktop
        let fixedWidth = 550;
        let fixedHeight = 160;
        let percentageWidth: number;
        let percentageHeight: number;

        percentageWidth = Math.round(100 * fixedWidth / screen.width);
        percentageHeight = Math.round(100 * fixedHeight / screen.height);

        dialogOptions = { width: percentageWidth, height: percentageHeight, displayInIframe: false }; // Desktop
    }
    return dialogOptions;
}

function receiveMessage(message: any, event: any) {
    addressesDialog.close();
    addressesDialog = null;

    if (message.message === "Send") {
        event.completed({ allowEvent: true });
    } else {
        event.completed({ allowEvent: false });
    }
}

function dialogClosed(event: any) {
    addressesDialog = null;
    event.completed({ allowEvent: false });
}