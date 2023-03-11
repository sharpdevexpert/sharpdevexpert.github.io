/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

/* eslint-env jquery */

var messageBanner;

var localStorageToken = "rfpninjatoken";
var loginEndPoint = "https://app.rfpninja.com/version-test/api/1.1/wf/remote-login";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("loginButton").onclick = loginFunc;
    document.getElementById("logoutButton").onclick = logoutFunc;

    handleLoginDivs();

    // Initialize the notification mechanism and hide it
    var element = document.querySelector(".MessageBanner");
    messageBanner = new components.MessageBanner(element);
    messageBanner.hideBanner();
  }
});

function loginFunc() {
  var username = $("#username").val();
  var password = $("#password").val();

  $.ajax({
    url: loginEndPoint,
    type: "POST",
    data: JSON.stringify({
      username: username,
      password: password,
    }),
    contentType: "application/json",
  })
    .done(function (data) {
      if (data.status == "success") {
        window.localStorage.setItem(localStorageToken, data.response.token);
        showNotification("Success", "Successfully logged in");
        handleLoginDivs();
      }
    })
    .fail(function (data) {
      if (data.responseJSON) {
        showNotification("Error", data.responseJSON.message);
      } else {
        showNotification("Error", data.responseText);
      }
    });
}

function logoutFunc() {
  $("#username").val("");
  $("#password").val("");

  window.localStorage.removeItem(localStorageToken);
  showNotification("Success", "Successfully logged out");
  handleLoginDivs();
}

function handleLoginDivs() {
  var loggedIn = window.localStorage.getItem(localStorageToken);

  if (loggedIn) {
    $("#logInDiv").hide();
    $("#logOutDiv").show();
  } else {
    $("#logInDiv").show();
    $("#logOutDiv").hide();
  }
}

// Helper function for displaying notifications
function showNotification(header, content) {
  $("#notification-header").text(header);
  $("#notification-body").text(content);
  messageBanner.showBanner();
  messageBanner.toggleExpansion();
}
