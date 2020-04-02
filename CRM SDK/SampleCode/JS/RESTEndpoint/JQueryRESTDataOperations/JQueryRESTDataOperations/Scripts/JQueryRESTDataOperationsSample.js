// =====================================================================
//  This file is part of the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================
// <snippetJQueryRESTDataOperations.JQueryRESTDataOperationsSample.js>
/// <reference path="SDK.JQuery.js" />
/// <reference path="jquery-1.9.1.js" />


var primaryContact = null;
var startButton;
var resetButton;
var output; //The <ol> element used by the writeMessage function to display text showing the progress of this sample.


$(function () {
 getFirstContactToBePrimaryContact();
 startButton = $("#start");
 startButton.click(createAccount);
 resetButton = $("#reset");
 resetButton.click(resetSample);
 output = $("#output");

});


function createAccount() {
 startButton.attr("disabled", "disabled");
 var account = {};
 account.Name = "Test Account Name";
 account.Description = "This account was created by the JQueryRESTDataOperations sample.";
 if (primaryContact != null) {
  //Set a lookup value
  writeMessage("Setting the primary contact to: " + primaryContact.FullName + ".");
  account.PrimaryContactId = { Id: primaryContact.ContactId, LogicalName: "contact", Name: primaryContact.FullName };

 }
 //Set a picklist value
 writeMessage("Setting Preferred Contact Method to E-mail.");
 account.PreferredContactMethodCode = { Value: 2 }; //E-mail

 //Set a money value
 writeMessage("Setting Annual Revenue to Two Million .");
 account.Revenue = { Value: "2000000.00" }; //Set Annual Revenue

 //Set a Boolean value
 writeMessage("Setting Contact Method Phone to \"Do Not Allow\".");
 account.DoNotPhone = true; //Do Not Allow

 //Add Two Tasks
 var today = new Date();
 var startDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 3); //Set a date three days in the future.

 var LowPriTask = { Subject: "Low Priority Task", ScheduledStart: startDate, PriorityCode: { Value: 0} }; //Low Priority Task
 var HighPriTask = { Subject: "High Priority Task", ScheduledStart: startDate, PriorityCode: { Value: 2} }; //High Priority Task
 account.Account_Tasks = [LowPriTask, HighPriTask]



 //Create the Account
 SDK.JQuery.createRecord(
     account,
     "Account",
     function (account) {
      writeMessage("The account named \"" + account.Name + "\" was created with the AccountId : \"" + account.AccountId + "\".");
      writeMessage("Retrieving account with the AccountId: \"" + account.AccountId + "\".");
      retrieveAccount(account.AccountId)
     },
     errorHandler
   );
     
}

function retrieveAccount(AccountId) {
 SDK.JQuery.retrieveRecord(
     AccountId,
     "Account",
     null, null,
     function (account) {
      writeMessage("Retrieved the account named \"" + account.Name + "\". This account was created on : \"" + account.CreatedOn + "\".");
      updateAccount(AccountId);
     },
     errorHandler
   );
}

function updateAccount(AccountId) {
 var account = {};
 writeMessage("Changing the account Name to \"Updated Account Name\".");
 account.Name = "Updated Account Name";
 writeMessage("Adding Address information");
 account.Address1_AddressTypeCode = { Value: 3 }; //Address 1: Address Type = Primary
 account.Address1_City = "Sammamish";
 account.Address1_Line1 = "123 Maple St.";
 account.Address1_PostalCode = "98074";
 account.Address1_StateOrProvince = "WA";
 writeMessage("Setting E-Mail address");
 account.EMailAddress1 = "someone@microsoft.com";


 SDK.JQuery.updateRecord(
     AccountId,
     account,
     "Account",
     function () {
      writeMessage("The account record changes were saved");
      deleteAccount(AccountId);
     },
     errorHandler
   );
}

function deleteAccount(AccountId) {
 if (confirm("Do you want to delete this account record?")) {
  writeMessage("You chose to delete the account record.");
  SDK.JQuery.deleteRecord(
       AccountId,
       "Account",
       function () {
        writeMessage("The account was deleted.");
        enableResetButton();
       },
       errorHandler
     );
 }
 else {
  var urlToAccountRecord = SDK.JQuery._getClientUrl() + "/main.aspx?etc=1&id=%7b" + AccountId + "%7d&pagetype=entityrecord";
  $("<li><span>You chose not to delete the record. You can view the record </span><a href='" +
  urlToAccountRecord + 
  "' target='_blank'>here</a></li>").appendTo(output);
  enableResetButton();
 }
}

function getFirstContactToBePrimaryContact() {

 SDK.JQuery.retrieveMultipleRecords(
     "Contact",
     "$select=ContactId,FullName&$top=1",
     function (results) {
      var firstResult = results[0];
      if (firstResult != null) {
       primaryContact = results[0];
      }
      else {
       writeMessage("No Contact records are available to set as the primary contact for the account.");
      }
     },
     errorHandler,
     function () {
      //OnComplete handler
     }
   );
}

function errorHandler(error) {
 writeMessage(error.message);
}

function enableResetButton() {
 resetButton.removeAttr("disabled");
}

function resetSample() {
 output.empty();
 startButton.removeAttr("disabled");
 resetButton.attr("disabled", "disabled");
}

//Helper function to write data to this page:
function writeMessage(message) {
 $("<li>" + message + "</li>").appendTo(output);
}

// </snippetJQueryRESTDataOperations.JQueryRESTDataOperationsSample.js>
