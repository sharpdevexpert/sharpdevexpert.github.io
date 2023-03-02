/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

/* eslint-env jquery */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
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

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;

var dialog;

var localStorageToken = "rfpninjatoken";

var proxyServer = "https://cors-anywhere.herokuapp.com/";
var endPoint = "https://app.rfpninja.com/version-test/api/1.1/wf/get-prompt-response";

function generate(event) {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      //write('Action failed. Error: ' + asyncResult.error.message);
    } else {
      var question = asyncResult.value;
      callService(question);
      openDialog();
    }
  });

  event.completed();
}

function callService(question) {
  var token = window.localStorage.getItem(localStorageToken);

  $.ajax({
    url: proxyServer + endPoint,
    type: "POST",
    data: JSON.stringify({
      prompt: question,
      format: "Mutiple Paragraphs",
      dataset: "Sales",
    }),
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
    },
  })
    .done(function (data) {
      try {
        dialog.close();
      } catch (err) { }

      if (data.status == "success") {
        if (!data.response["returned-an-error"]) {
          var prompt = "";

          if (!data.response.prompt) {
            prompt = data.response.prompt;
          }

          Office.context.document.setSelectedDataAsync(prompt + data.response.response, function (asyncResult) {
            if (asyncResult.status === "failed") {
              // Show error message.
            } else {
              // Show success message.
            }
          });
        } else {
          var message = data.response["error-status-message"];

          var messageToDialog = JSON.stringify({
            name: message,
          });

          openDialog();
          dialog.messageChild(messageToDialog);
        }
      }
    })
    .fail(function (data) {
      var message = data.responseText;

      if (data.responseJSON) {
        message = data.responseJSON.translation;
      }

      var messageToDialog = JSON.stringify({
        name: message,
      });

      openDialog();
      dialog.messageChild(messageToDialog);
    });
}

function openDialog() {
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/dialog.html",
    { height: 10, width: 15, displayInIframe: true },
    function (asyncResult) {
      dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
  );
}

function processMessage(arg) {
  dialog.close();
}

Office.actions.associate("generate", generate);
