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
      if (data.status == "success") {
        Office.context.document.setSelectedDataAsync(
          data.response.prompt + data.response.response + "\n",
          function (asyncResult) {
          if (asyncResult.status === "failed") {
            // Show error message.
          } else {
            // Show success message.
          }
        });
      }
    })
    .fail(function (data) {
      return JSON.stringify(data);
    });
}

function openDialog() {
  Office.context.ui.displayDialogAsync("https://localhost:3000/src/dialog.html", { displayInIframe: true }, null);
}

Office.actions.associate("generate", generate);
