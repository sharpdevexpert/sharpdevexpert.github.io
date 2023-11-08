/* global Office */

/* eslint-env jquery */

(function () {
  "use strict";

  Office.onReady().then(function () {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onMessageFromParent);
  });

  function onMessageFromParent(arg) {
    const messageFromParent = JSON.parse(arg.message);
    $("p").text(messageFromParent.name);
  }
});
