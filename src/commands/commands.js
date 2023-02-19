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

var proxyServer = "https://cors-anywhere.herokuapp.com/";
var endPoint = "https://api.sustainably.ai/qaquery";

function generate(event) {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      //write('Action failed. Error: ' + asyncResult.error.message);
    } else {
      var question = asyncResult.value;
      callService(question);
    }
  });

  event.completed();
}

function callService(question) {
  $.ajax({
    url: proxyServer + "https://api.sustainably.ai/qaquery",
    type: "POST",
    data: JSON.stringify({
      "prompt": question,
      "API_KEY": "H888biE9ADWVYSKrmU7c53Yrv",
      "response-format": "Mutiple Paragraphs",
      "data-set": "Sales",
    }),
    contentType: "application/json",
  })
    .done(function (data) {
      //return data.qaresponse;
      Office.context.document.setSelectedDataAsync(data.prompt + data.qaresponse + "\n", function (asyncResult) {
        if (asyncResult.status === "failed") {
          // Show error message.
        } else {
          // Show success message.
        }
      });
    })
    .fail(function (status) {
      return JSON.stringify(status);
    });
}

Office.actions.associate("generate", generate);
