
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

/// <reference path="jquery-1.9.1.js" />

//GetGlobalContext function exists in ClientGlobalContext.js.aspx so the
//host HTML page must have a reference to ClientGlobalContext.js.aspx.
var context = GetGlobalContext();

//Retrieve the client url
var clientUrl = context.getClientUrl();



//The XRM OData end-point
var ODATA_ENDPOINT = "/XRMServices/2011/OrganizationData.svc";



function createRecord(entityObject, odataSetName, successCallback, errorCallback)
{
 ///	<summary>
 ///		Uses jQuery's AJAX object to call the Microsoft Dynamics CRM OData endpoint to
 ///     Create a new record
 ///	</summary>
 /// <param name="entityObject" type="Object" required="true">
 ///		1: entity - a loose-type object representing an OData entity. any fields
 ///                 on this object must be camel-cased and named exactly as they 
 ///                 appear in entity metadata
 ///	</param>
 /// <param name="odataSetName" type="string" required="true">
 ///		1: set -    a string representing an OData Set. OData provides uri access
 ///                 to any CRM entity collection. examples: AccountSet, ContactSet,
 ///                 OpportunitySet. 
 ///	</param>
 /// <param name="successCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon success
 ///                 of the ajax invocation.
 ///	</param>
 /// <param name="errorCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon error
 ///                 of the ajax invocation.
 ///	</param>


 //entityObject is required
 if (!entityObject)
 {
  alert("entityObject is required.");
  return;
 }
 //odataSetName is required, i.e. "AccountSet"
 if (!odataSetName)
 {
  alert("odataSetName is required.");
  return;
 }
 else
 { odataSetName = encodeURIComponent(odataSetName); }

 //Parse the entity object into JSON
 var jsonEntity = window.JSON.stringify(entityObject);

 //Asynchronous AJAX function to Create a CRM record using OData
 $.ajax({
  type: "POST",
  contentType: "application/json; charset=utf-8",
  datatype: "json",
  url: clientUrl + ODATA_ENDPOINT + "/" + odataSetName,
  data: jsonEntity,
  beforeSend: function (XMLHttpRequest)
  {
   //Specifying this header ensures that the results will be returned as JSON.             
   XMLHttpRequest.setRequestHeader("Accept", "application/json");
  },
  success: function (data, textStatus, XmlHttpRequest)
  {
   if (successCallback)
   {
    successCallback(data.d, textStatus, XmlHttpRequest);
   }
  },
  error: function (XmlHttpRequest, textStatus, errorThrown)
  {
   if (errorCallback)
    errorCallback(XmlHttpRequest, textStatus, errorThrown);
   else
    errorHandler(XmlHttpRequest, textStatus, errorThrown);
  }
 });
}

function retrieveRecord(id, odataSetName, successCallback, errorCallback)
{
 ///	<summary>
 ///		Uses jQuery's AJAX object to call the Microsoft Dynamics CRM OData endpoint to
 ///     retrieve an existing record
 ///	</summary>
 /// <param name="id" type="guid" required="true">
 ///		1: id -     the guid (primarykey) of the record to be retrieved
 ///	</param>
 /// <param name="odataSetName" type="string" required="true">
 ///		1: set -    a string representing an OData Set. OData provides uri access
 ///                 to any CRM entity collection. examples: AccountSet, ContactSet,
 ///                 OpportunitySet. 
 ///	</param>
 /// <param name="successCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon success
 ///                 of the ajax invocation.
 ///	</param>
 /// <param name="errorCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon error
 ///                 of the ajax invocation.
 ///	</param>

 //id is required
 if (!id)
 {
  alert("record id is required.");
  return;
 }
 else
 {
  id = encodeURIComponent(id);
 }
 //odataSetName is required, i.e. "AccountSet"
 if (!odataSetName)
 {
  alert("odataSetName is required.");
  return;
 }
 else
 { odataSetName = encodeURIComponent(odataSetName); }

 //Asynchronous AJAX function to Retrieve a CRM record using OData
 $.ajax({
  type: "GET",
  contentType: "application/json; charset=utf-8",
  datatype: "json",
  url: clientUrl + ODATA_ENDPOINT + "/" + odataSetName + "(guid'" + id + "')",
  beforeSend: function (XMLHttpRequest)
  {
   //Specifying this header ensures that the results will be returned as JSON.             
   XMLHttpRequest.setRequestHeader("Accept", "application/json");
  },
  success: function (data, textStatus, XmlHttpRequest)
  {
   if (successCallback)
   {
    successCallback(data.d, textStatus, XmlHttpRequest);
   }
  },
  error: function (XmlHttpRequest, textStatus, errorThrown)
  {
   if (errorCallback)
    errorCallback(XmlHttpRequest, textStatus, errorThrown);
   else
    errorHandler(XmlHttpRequest, textStatus, errorThrown);
  }
 });
}

function retrieveMultiple(odataSetName, filter, successCallback, errorCallback)
{
 ///	<summary>
 ///		Uses jQuery's AJAX object to call the Microsoft Dynamics CRM OData endpoint to
 ///     Retrieve multiple records
 ///	</summary>
 /// <param name="odataSetName" type="string" required="true">
 ///		1: set -    a string representing an OData Set. OData provides uri access
 ///                 to any CRM entity collection. examples: AccountSet, ContactSet,
 ///                 OpportunitySet. 
 ///	</param>
 /// <param name="filter" type="string">
 ///		1: filter - a string representing the filter that is appended to the odatasetname
 ///                 of the OData URI.     
 ///	</param>
 /// <param name="successCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon success
 ///                 of the ajax invocation.
 ///	</param>
 /// <param name="errorCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon error
 ///                 of the ajax invocation.
 ///	</param>
 //odataSetName is required, i.e. "AccountSet"
 if (!odataSetName)
 {
  alert("odataSetName is required.");
  return;
 }
 else
 { odataSetName = encodeURIComponent(odataSetName); }

 //Build the URI
 var odataUri = clientUrl + ODATA_ENDPOINT + "/" + odataSetName;

 //If a filter is supplied, append it to the OData URI
 if (filter)
 {
  odataUri += "?$filter=" + encodeURIComponent(filter);
 }

 //Asynchronous AJAX function to Retrieve CRM records using OData
 $.ajax({
  type: "GET",
  contentType: "application/json; charset=utf-8",
  datatype: "json",
  url: odataUri,
  beforeSend: function (XMLHttpRequest)
  {
   //Specifying this header ensures that the results will be returned as JSON.             
   XMLHttpRequest.setRequestHeader("Accept", "application/json");
  },
  success: function (data, textStatus, XmlHttpRequest)
  {
   if (successCallback)
   {
    if (data && data.d && data.d.results)
    {
     successCallback(data.d.results, textStatus, XmlHttpRequest);
    }
    else if (data && data.d)
    {
     successCallback(data.d, textStatus, XmlHttpRequest);
    }
    else
    {
     successCallback(data, textStatus, XmlHttpRequest);
    }
   }
  },
  error: function (XmlHttpRequest, textStatus, errorThrown)
  {
   if (errorCallback)
    errorCallback(XmlHttpRequest, textStatus, errorThrown);
   else
    errorHandler(XmlHttpRequest, textStatus, errorThrown);
  }
 });
}

function updateRecord(id, entityObject, odataSetName, successCallback, errorCallback)
{
 ///	<summary>
 ///		Uses jQuery's AJAX object to call the Microsoft Dynamics CRM OData endpoint to
 ///     update an existing record
 ///	</summary>
 /// <param name="id" type="guid" required="true">
 ///		1: id -     the guid (primarykey) of the record to be retrieved
 ///	</param>
 /// <param name="entityObject" type="Object" required="true">
 ///		1: entity - a loose-type object representing an OData entity. any fields
 ///                 on this object must be camel-cased and named exactly as they 
 ///                 appear in entity metadata
 ///	</param>
 /// <param name="odataSetName" type="string" required="true">
 ///		1: set -    a string representing an OData Set. OData provides uri access
 ///                 to any CRM entity collection. examples: AccountSet, ContactSet,
 ///                 OpportunitySet. 
 ///	</param>
 /// <param name="successCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon success
 ///                 of the ajax invocation.
 ///	</param>
 /// <param name="errorCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon error
 ///                 of the ajax invocation.
 ///	</param>

 //id is required
 if (!id)
 {
  alert("record id is required.");
  return;
 }
 else
 { id = encodeURIComponent(id); }
 //odataSetName is required, i.e. "AccountSet"
 if (!odataSetName)
 {
  alert("odataSetName is required.");
  return;
 }
 else
 { odataSetName = encodeURIComponent(odataSetName); }

 if (!entityObject)
 {
  alert("entityObject is required.");
  return;
 }

 //Parse the entity object into JSON
 var jsonEntity = window.JSON.stringify(entityObject);

 //Asynchronous AJAX function to Update a CRM record using OData
 $.ajax({
  type: "POST",
  contentType: "application/json; charset=utf-8",
  datatype: "json",
  data: jsonEntity,
  url: clientUrl + ODATA_ENDPOINT + "/" + odataSetName + "(guid'" + id + "')",
  beforeSend: function (XMLHttpRequest)
  {
   //Specifying this header ensures that the results will be returned as JSON.             
   XMLHttpRequest.setRequestHeader("Accept", "application/json");

   //Specify the HTTP method MERGE to update just the changes you are submitting.             
   XMLHttpRequest.setRequestHeader("X-HTTP-Method", "MERGE");
  },
  success: function (data, textStatus, XmlHttpRequest)
  {
   //The MERGE does not return any data at all, so we'll add the id 
   //onto the data object so it can be leveraged in a Callback. When data 
   //is used in the callback function, the field will be named generically, "id"
   data = new Object();
   data.id = id;
   if (successCallback)
   {
    successCallback(data, textStatus, XmlHttpRequest);
   }
  },
  error: function (XmlHttpRequest, textStatus, errorThrown)
  {
   if (errorCallback)
    errorCallback(XmlHttpRequest, textStatus, errorThrown);
   else
    errorHandler(XmlHttpRequest, textStatus, errorThrown);
  }
 });
}

function deleteRecord(id, odataSetName, successCallback, errorCallback)
{
 ///	<summary>
 ///		Uses jQuery's AJAX object to call the Microsoft Dynamics CRM OData endpoint to
 ///     delete an existing record
 ///	</summary>
 /// <param name="id" type="guid" required="true">
 ///		1: id -     the guid (primarykey) of the record to be retrieved
 ///	</param>
 /// <param name="odataSetName" type="string" required="true">
 ///		1: set -    a string representing an OData Set. OData provides uri access
 ///                 to any CRM entity collection. examples: AccountSet, ContactSet,
 ///                 OpportunitySet. 
 ///	</param>
 /// <param name="successCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon success
 ///                 of the ajax invocation.
 ///	</param>
 /// <param name="errorCallback" type="function" >
 ///		1: callback-a function that can be supplied as a callback upon error
 ///                 of the ajax invocation.
 ///	</param>

 //id is required
 if (!id)
 {
  alert("record id is required.");
  return;
 }
 else
 { id = encodeURIComponent(id); }

 //odataSetName is required, i.e. "AccountSet"
 if (!odataSetName)
 {
  alert("odataSetName is required.");
  return;
 }
 else
 { odataSetName = encodeURIComponent(odataSetName); }

 //Asynchronous AJAX function to Delete a CRM record using OData
 $.ajax({
  type: "POST",
  contentType: "application/json; charset=utf-8",
  datatype: "json",
  url: clientUrl + ODATA_ENDPOINT + "/" + odataSetName + "(guid'" + id + "')",
  beforeSend: function (XMLHttpRequest)
  {
   //Specifying this header ensures that the results will be returned as JSON.                 
   XMLHttpRequest.setRequestHeader("Accept", "application/json");

   //Specify the HTTP method DELETE to perform a delete operation.                 
   XMLHttpRequest.setRequestHeader("X-HTTP-Method", "DELETE");
  },
  success: function (data, textStatus, XmlHttpRequest)
  {
   if (successCallback)
   {
     successCallback(null, textStatus, XmlHttpRequest);
   }
  },
  error: function (XmlHttpRequest, textStatus, errorThrown)
  {
   if (errorCallback)
    errorCallback(XmlHttpRequest, textStatus, errorThrown);
   else
    errorHandler(XmlHttpRequest, textStatus, errorThrown);
  }
 });
}

///	<summary>
///		A function that will display the error results of an AJAX operation
///	</summary>
function errorHandler(xmlHttpRequest, textStatus, errorThrown)
{
 alert("Error : " + textStatus + ": " + xmlHttpRequest.statusText);
}