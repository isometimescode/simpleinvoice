/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 * This is a copy of the example at:
 * https://developers.google.com/apps-script/guides/dialogs#file-open_dialogs
 *
 * License: MIT
 *
 * Copyright 2016 Toni Wells <isometimescode.com>
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invoicing')
      .addItem('Create New Invoice', 'showPicker')
      .addToUi();
}

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Choose Google Docs Template File');
}

/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}