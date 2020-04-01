// =====================================================================
//
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
//
// =====================================================================
//<snippetJavaScriptRESTAssociateDisassociateJS>
/// <reference path="SDK.REST.js" />
var btnStartAssociateAccounts,
btnStartAssociateAccountToEmail,
btnDisassociateAccounts,
btnResetSample,
parentAccount,
childAccount,
accountRelatedToEmail,
emailRecord,
output, alertFlag;

document.onreadystatechange = function () {
 ///<summary>
 /// Initializes the sample when the document is ready
 ///</summary>
 if (document.readyState == "complete") {

  btnStartAssociateAccounts = document.getElementById("btnStartAssociateAccounts");
  btnStartAssociateAccountToEmail = document.getElementById("btnStartAssociateAccountToEmail");
  btnResetSample = document.getElementById("btnResetSample");
  alertFlag = document.getElementById("dispalert");
  output = document.getElementById("output");

  btnStartAssociateAccounts.onclick = startAssociateAccounts;
  btnStartAssociateAccountToEmail.onclick = startAssociateAccountToEmail;
  btnResetSample.onclick = resetSample;
 }
}

function startAssociateAccounts() {
 ///<summary>
 /// Creates two account records asynchronously and displays a message showing the Id values of the created records
 /// Then calls the associateAccounts function to associate them.
 ///</summary>
 disableElement(btnStartAssociateAccounts);
 disableElement(btnStartAssociateAccountToEmail);
 //Create the first account record
 SDK.REST.createRecord({ Name: "Parent Account", Description: "This Account will be the parent account" },
	"Account",
	function (account) {
	 parentAccount = account;
	 writeMessage("Account Id: {" + account.AccountId + "} Name: \"" + account.Name + "\" created as the parent account.");
	 // Create the second account record
	 SDK.REST.createRecord({ Name: "Child Account", Description: "This Account will be the child account" },
		"Account",
		function (account) {
		 childAccount = account;
		 writeMessage("Account Id: {" + account.AccountId + "} Name: \"" + account.Name + "\" created as the child account.");
		 // Associate the accounts that were created.
		 associateAccounts(parentAccount, childAccount);
		},
	 errorHandler);
	},
	 errorHandler);

 scrollToResults();
}

function associateAccounts(parentAccount, childAccount) {
 ///<summary>
 /// Associates two account records asynchronously and displays messages so you can verify the association
 /// Displays  buttons to disassociate or delete the accounts.
 /// Enables the reset button
 ///</summary>
 SDK.REST.associateRecords(parentAccount.AccountId,
	"Account",
	"Referencedaccount_parent_account",
	childAccount.AccountId,
	"Account",
	function () {
	 writeMessage("Association successful.");
	 showLinksToVerifyAccountAssociation();
	 showButtonToDisassociateAccounts();
	 showButtonToDeleteAccounts();
	 enableElement(btnResetSample);
	 if (alertFlag.checked == true)
	  alert("Association successful.");
	},
	errorHandler);
}

function getUrl()
{
	var context;
	var url;
	if (typeof GetGlobalContext != "undefined")
	{ context = GetGlobalContext(); }
	else
	{
		if ((typeof Xrm != "undefined") && (typeof Xrm.Page != "undefined") && (typeof Xrm.Page.context != "undefined"))
		{
			context = Xrm.Page.context;
		}
		else
		{ throw new Error("Context is not available."); }
	}


		url = context.getClientUrl();


	return url;
}

function showLinksToVerifyAccountAssociation() {
 ///<summary>
 /// Displays links to verify the associated accounts.
 ///</summary>
 var message = document.createElement("div");
 var message1 = document.createElement("span");
 setText(message, "Open the ");
 var childRecordlink = document.createElement("a");
 childRecordlink.href = getUrl() + "/main.aspx?etn=account&pagetype=entityrecord&id=%7B" + childAccount.AccountId + "%7D";
 childRecordlink.target = "_blank";
 setText(childRecordlink, childAccount.Name + " record");
 var message2 = document.createElement("span");
 setText(message2, " to verify that the Parent Account field is set to the ");
 var parentRecordlink = document.createElement("a");
 parentRecordlink.href = getUrl() + "/main.aspx?etn=account&pagetype=entityrecord&id=%7B" + parentAccount.AccountId + "%7D";
 parentRecordlink.target = "_blank";
 setText(parentRecordlink, parentAccount.Name + " record");
 period = document.createElement("span");
 setText(period, ".");
 message.appendChild(message1);
 message.appendChild(childRecordlink);
 message.appendChild(message2);
 message.appendChild(parentRecordlink);
 message.appendChild(period);
 writeMessage(message);
}

function showLinksToVerifyDisAssociation() {
 ///<summary>
 /// Displays links to verify accounts were disassociated
 ///</summary>
 var message = document.createElement("div");
 var message1 = document.createElement("span");
 setText(message, "Open the ");
 var childRecordlink = document.createElement("a");
 childRecordlink.href = getUrl() + "/main.aspx?etn=account&pagetype=entityrecord&id=%7B" + childAccount.AccountId + "%7D";
 childRecordlink.target = "_blank";
 setText(childRecordlink, childAccount.Name + " record");
 var message2 = document.createElement("span");
 setText(message2, " to verify that the Parent Account field not set to the ");
 var parentRecordlink = document.createElement("a");
 parentRecordlink.href = getUrl() + "/main.aspx?etn=account&pagetype=entityrecord&id=%7B" + parentAccount.AccountId + "%7D";
 parentRecordlink.target = "_blank";
 setText(parentRecordlink, parentAccount.Name + " record");
 period = document.createElement("span");
 setText(period, ".");
 message.appendChild(message1);
 message.appendChild(childRecordlink);
 message.appendChild(message2);
 message.appendChild(parentRecordlink);
 message.appendChild(period);
 writeMessage(message);
}

function showButtonToDeleteAccounts() {
 ///<summary>
 /// Diplays a button to allow for deletion of accounts created for this sample.
 ///</summary>
 var btnDeleteRecords = document.createElement("button");
 setText(btnDeleteRecords, "Delete both records");
 btnDeleteRecords.title = "Delete both records";
 btnDeleteRecords.onclick = function () {
  disableElement(btnDisassociateAccounts);
  disableElement(this);
  SDK.REST.deleteRecord(childAccount.AccountId,
			"Account",
			function () { writeMessage(childAccount.Name + " record deleted."); },
			errorHandler);
  SDK.REST.deleteRecord(parentAccount.AccountId,
			"Account",
			function () {
			 writeMessage(parentAccount.Name + " record deleted.");
			 if (alertFlag.checked == true)
			  alert("Records deleted.");
			},
			errorHandler);
 }
 writeMessage(btnDeleteRecords);
}

function showButtonToDisassociateAccounts() {
 ///<summary>
 /// Displays a button to allow for disassociation of accounts associated by this sample.
 /// Then shows links to verify disassociation
 ///</summary>
 btnDisassociateAccounts = document.createElement("button");
 setText(btnDisassociateAccounts, "Disassociate the records");
 btnDisassociateAccounts.title = "Disassociate the records";
 btnDisassociateAccounts.onclick = function () {
  SDK.REST.disassociateRecords(parentAccount.AccountId,
	"Account",
	"Referencedaccount_parent_account",
	childAccount.AccountId,
	function () {
	 showLinksToVerifyDisAssociation();
	 if (alertFlag.checked == true)
	  alert("Disassociation successful.");
	},
	errorHandler);
 }
 writeMessage(btnDisassociateAccounts);
}

function startAssociateAccountToEmail() {
 ///<summary>
 /// Creates an account and email record and displays message to verify that they were created
 /// Then calls associateAccountAsActivityPartyToEmail to associate them
 ///</summary>
 disableElement(btnStartAssociateAccounts);
 disableElement(btnStartAssociateAccountToEmail);

 //Create one account and one email record.
 SDK.REST.createRecord({ Name: "Email Sender Account", Description: "This Account will be the account shown as the sender of an email." },
	"Account",
	function (account) {
	 accountRelatedToEmail = account;
	 writeMessage("Account Id: {" + accountRelatedToEmail.AccountId + "} Name: \"" + accountRelatedToEmail.Name + "\" created as an account to be the sender of an email.");
	 SDK.REST.createRecord({ Subject: "Email Activity", Description: "This email will be shown as sent by the " + accountRelatedToEmail.Name + " record." },
		"Email",
		function (email) {
		 emailRecord = email;
		 writeMessage("Email ActivityId: {" + emailRecord.ActivityId + "} Subject: \"" + emailRecord.Subject + "\" created to be shown as sent from the " + accountRelatedToEmail.Name + " record.");
		 associateAccountAsActivityPartyToEmail();
		},
	 errorHandler);
	},
	 errorHandler);
 scrollToResults();

}

function associateAccountAsActivityPartyToEmail() {
 ///<summary>
 /// Creates a new ActivityParty record to associate an account record to an email record, making the account the sender of the email.
 /// Shows links to verify the association
 /// Shows buttons to delete the records. It is not possible to disassociate the records by deleting the ActivityParty record using the REST endpoint.
 ///</summary>
 var activityParty = {
  PartyId:
 {
  Id: accountRelatedToEmail.AccountId,
  LogicalName: "account"
 },
  ActivityId: {
   Id: emailRecord.ActivityId,
   LogicalName: "email"
  },
  // Set the participation type (what role the party has on the activity). For this
  // example, we'll put the account in the From field (which has a value of 1).
  // See http://msdn.microsoft.com/en-us/library/gg328549.aspx for other options.
  ParticipationTypeMask: { Value: 1 }
 };

 SDK.REST.createRecord(activityParty,
 "ActivityParty",
 function (ap) {
  test_activityPartyId = ap.ActivityPartyId;
  writeMessage("Created new ActivityParty ActivityPartyId: {" + ap.ActivityPartyId + "}. The account is now related to the email.");
  showLinksToVerifyActivityPartyAssociation();
  showButtonToDeleteAccountAndEmailRecords();
  enableElement(btnResetSample);
  if (alertFlag.checked == true)
   alert("The account is now related to the email.");
 },
 errorHandler);
}

// Note: Attempting to disassociate an activityparty relationship by deleting the ActivityParty record is not allowed.

function showLinksToVerifyActivityPartyAssociation() {
 ///<summary>
 /// Shows links to verify that an account and email are linked
 ///</summary>
 var message = document.createElement("div");
 var message1 = document.createElement("span");
 setText(message1, "Open the ");
 var childRecordlink = document.createElement("a");
 childRecordlink.href = getUrl() + "/main.aspx?etn=email&pagetype=entityrecord&id=%7B" + emailRecord.ActivityId + "%7D";
 childRecordlink.target = "_blank";
 setText(childRecordlink, emailRecord.Subject + " record");
 var message2 = document.createElement("span");
 setText(message2, " to verify that the ");
 var parentRecordlink = document.createElement("a");
 parentRecordlink.href = getUrl() + "/main.aspx?etn=account&pagetype=entityrecord&id=%7B" + accountRelatedToEmail.AccountId + "%7D";
 parentRecordlink.target = "_blank";
 setText(parentRecordlink, accountRelatedToEmail.Name + " record");
 message3 = document.createElement("span");
 setText(message3, " is set as the sender.");
 message.appendChild(message1);
 message.appendChild(childRecordlink);
 message.appendChild(message2);
 message.appendChild(parentRecordlink);
 message.appendChild(message3);
 writeMessage(message);
}

function showButtonToDeleteAccountAndEmailRecords() {
 ///<summary>
 /// Shows button to delete the account and email records created for this sample
 ///</summary>
 var btnDeleteRecords = document.createElement("button");
 setText(btnDeleteRecords, "Delete both records");
 btnDeleteRecords.title = "Delete both records";
 btnDeleteRecords.onclick = function () {
  disableElement(this);
  SDK.REST.deleteRecord(accountRelatedToEmail.AccountId,
			"Account",
			function () { writeMessage(accountRelatedToEmail.Name + " record deleted."); },
			errorHandler);
  SDK.REST.deleteRecord(emailRecord.ActivityId,
			"Email",
			function () {
			 writeMessage(emailRecord.Subject + " record deleted.");
			 if (alertFlag.checked == true)
			  alert("Records deleted.");
			},
			errorHandler);
 }
 writeMessage(btnDeleteRecords);
}

function errorHandler(error) {
 ///<summary>
 /// Displays the message property of errors
 ///</summary>
 writeMessage(error.message);
 if (alertFlag.checked == true)
  alert(error.message);
}

function enableElement(element) {
 ///<summary>
 /// Enables an element that is disabled.
 ///</summary>
 element.removeAttribute("disabled");
}

function disableElement(element) {
 ///<summary>
 /// Disables an element that is enabled.
 ///</summary>
 element.setAttribute("disabled", "disabled");
}

function resetSample() {
 ///<summary>
 /// Clears out the results area and enable buttons to start one of the samples again.
 ///</summary>
 output.innerHTML = "";
 enableElement(btnStartAssociateAccounts);
 enableElement(btnStartAssociateAccountToEmail)
 disableElement(btnResetSample);
 if (alertFlag.checked == true)
  alert("Reset complete.");
}

//Helper function to write data to this page:
function writeMessage(message) {
 ///<summary>
 /// Displays a message or appends an element to the results area.
 ///</summary>
 var li = document.createElement("li");
 if (typeof (message) == "string") {

  setText(li, message);

 }
 else {
  li.appendChild(message);
 }

 output.appendChild(li);
}

function scrollToResults() {
 ///<summary>
 /// Scrolls to the bottom of the page so the results area can be seen.
 ///</summary>
 window.scrollTo(0, document.body.scrollHeight);
}

function setText(node, text) {
 ///<summary>
 /// Used to set the text content of elements to manage differences between browsers.
 ///</summary>
 if (typeof (node.innerText) != "undefined") {
  node.innerText = text;
 }
 else {
  node.textContent = text;
 }
}
//</snippetJavaScriptRESTAssociateDisassociateJS>