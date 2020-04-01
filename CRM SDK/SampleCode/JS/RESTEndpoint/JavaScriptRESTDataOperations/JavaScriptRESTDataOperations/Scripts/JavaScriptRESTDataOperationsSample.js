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

/// <reference path="SDK.REST.js" />

//<snippetJavaScriptRESTDataOperationsSampleJS>
var primaryContact = null;
var startButton;
var resetButton;
var output; //The <ol> element used by the writeMessage function to display text showing the progress of this sample.

document.onreadystatechange = function () {
 if (document.readyState == "complete") {
  getFirstContactToBePrimaryContact();

  startButton = document.getElementById("start");
  resetButton = document.getElementById("reset");
  output = document.getElementById("output");

  startButton.onclick = createAccount;
  resetButton.onclick = resetSample;
 }
}

function createAccount() {
 var account = {};
 account.Name = "Test Account Name";
 account.Description = "This account was created by the JavaScriptRESTDataOperations sample.";
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
 SDK.REST.createRecord(
     account,
     "Account",
     function (account) {
      writeMessage("The account named \"" + account.Name + "\" was created with the AccountId : \"" + account.AccountId + "\".");
      writeMessage("Retrieving account with the AccountId: \"" + account.AccountId + "\".");
      retrieveAccount(account.AccountId)
     },
     errorHandler
   );
 this.setAttribute("disabled", "disabled");
}

function retrieveAccount(AccountId) {
 SDK.REST.retrieveRecord(
     AccountId,
     "Account",
     null,null,
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


 SDK.REST.updateRecord(
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
  SDK.REST.deleteRecord(
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
  var li = document.createElement("li");

  var span = document.createElement("span");

  setElementText(span, "You chose not to delete the record. You can view the record ");
		

  var link = document.createElement("a");
  link.href = SDK.REST._getClientUrl() + "/main.aspx?etc=1&id=%7b" + AccountId + "%7d&pagetype=entityrecord";
  link.target = "_blank";
  setElementText(link, "here");


  li.appendChild(span);
  li.appendChild(link);
  output.appendChild(li);
  enableResetButton();

 }
}

function getFirstContactToBePrimaryContact() {

 SDK.REST.retrieveMultipleRecords(
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
 resetButton.removeAttribute("disabled");
}

function resetSample() {
 output.innerHTML = "";
 startButton.removeAttribute("disabled");
 resetButton.setAttribute("disabled", "disabled");
}

//Helper function to write data to this page:
function writeMessage(message) {
	var li = document.createElement("li");

	setElementText(li, message);


 output.appendChild(li);
}
//Because Firefox doesn't  support innerText
function setElementText(element, text)
{
	if (element.innerText != undefined)
	{
		element.innerText = text;
	}
	else
	{
		element.textContent = text;
	}
}

//</snippetJavaScriptRESTDataOperationsSampleJS>