/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

/* eslint-env jquery */

var messageBanner;

var localStorageToken = "rfpninjatoken";
var localStorageDataSet = "rfpninjadataset";
var localStorageDataSetCollection = "rfpninjadatasetCollection";

var loginEndPoint = "https://app.rfpninja.com/version-test/api/1.1/wf/remote-login";
var dataSetsEndPoint = "https://app.rfpninja.com/version-test/api/1.1/wf/get-datasets";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Assign events to DOM elements
    document.getElementById("loginButton").onclick = loginFunc;
    document.getElementById("logoutButton").onclick = logoutFunc;
    document.getElementById("datasets").onchange = onSelectDataSet;

    // Handle divs based on login status
    handleLoginDivs();

    // Populate dropdown after refresh
    var dataSetsCollection = window.localStorage.getItem(localStorageDataSetCollection);

    if (dataSetsCollection) {
      dataSetsDropDown(JSON.parse(dataSetsCollection));
      document.getElementById("datasets").value = window.localStorage.getItem(localStorageDataSet);
    }

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

        getDataSets();
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
  window.localStorage.removeItem(localStorageDataSet);
  window.localStorage.removeItem(localStorageDataSetCollection);

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

function getDataSets() {
  var token = window.localStorage.getItem(localStorageToken);

  $.ajax({
    url: dataSetsEndPoint,
    type: "GET",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
    },
  })
    .done(function (data) {
      if (data.status == "success") {
        var dataSets = data.response.datasets;

        dataSetsDropDown(dataSets);

        // Default dataset is first one
        if (dataSets.length > 0) {
          window.localStorage.setItem(localStorageDataSet, dataSets[0]);
        }

        window.localStorage.setItem(localStorageDataSetCollection, JSON.stringify(dataSets));

        showNotification("Success", "Successfully logged in");
        handleLoginDivs();
      }
    })
    .fail(function (data) {});
}

function dataSetsDropDown(dataSets) {
  var select = document.getElementById("datasets");
  $(select).empty();

  dataSets.forEach((element) => {
    var opt = document.createElement("option");
    opt.value = element;
    opt.innerHTML = element;
    select.appendChild(opt);
  });
}

function onSelectDataSet() {
  var selection = document.getElementById("datasets").value;
  window.localStorage.setItem(localStorageDataSet, selection);

  showNotification("Success", "Choosen DataSet is: " + selection);
}
