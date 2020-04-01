// =====================================================================
//
//  This file is part of the Microsoft Dynamics CRM SDK Code Samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or online documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
//
// =====================================================================

//<snippetModernSoapApp2>
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace ModernSoapApp
{
    public static class HttpRequestBuilder
    {
        /// <summary>
        /// Retrieve entity record data from the organization web service. 
        /// </summary>
        /// <param name="accessToken">The web service authentication access token.</param>
        /// <param name="Columns">The entity attributes to retrieve.</param>
        /// <param name="entity">The target entity for which the data should be retreived.</param>
        /// <returns>Response from the web service.</returns>
        /// <remarks>Builds a SOAP HTTP request using passed parameters and sends the request to the server.</remarks>
        public static async Task<string> RetrieveMultiple(string accessToken, string[] Columns, string entity)
        {
            // Build a list of entity attributes to retrieve as a string.
            string columnsSet = string.Empty;
            foreach (string Column in Columns)
            {
                columnsSet += "<b:string>" + Column + "</b:string>";
            }

            // Default SOAP envelope string. This XML code was obtained using the SOAPLogger tool.
            string xmlSOAP =
             @"<s:Envelope xmlns:s='http://schemas.xmlsoap.org/soap/envelope/'>
                <s:Body>
                  <RetrieveMultiple xmlns='http://schemas.microsoft.com/xrm/2011/Contracts/Services' xmlns:i='http://www.w3.org/2001/XMLSchema-instance'>
                    <query i:type='a:QueryExpression' xmlns:a='http://schemas.microsoft.com/xrm/2011/Contracts'><a:ColumnSet>
                    <a:AllColumns>false</a:AllColumns><a:Columns xmlns:b='http://schemas.microsoft.com/2003/10/Serialization/Arrays'>" + columnsSet +
                   @"</a:Columns></a:ColumnSet><a:Criteria><a:Conditions /><a:FilterOperator>And</a:FilterOperator><a:Filters /></a:Criteria>
                    <a:Distinct>false</a:Distinct><a:EntityName>" + entity + @"</a:EntityName><a:LinkEntities /><a:Orders />
                    <a:PageInfo><a:Count>0</a:Count><a:PageNumber>0</a:PageNumber><a:PagingCookie i:nil='true' />
                    <a:ReturnTotalRecordCount>false</a:ReturnTotalRecordCount>
                    </a:PageInfo><a:NoLock>false</a:NoLock></query>
                  </RetrieveMultiple>
                </s:Body>
              </s:Envelope>";

            // The URL for the SOAP endpoint of the organization web service.
            string url = CurrentEnvironment.CrmServiceUrl + "/XRMServices/2011/Organization.svc/web";

            // Use the RetrieveMultiple CRM message as the SOAP action.
            string SOAPAction = "http://schemas.microsoft.com/xrm/2011/Contracts/Services/IOrganizationService/RetrieveMultiple";

            // Create a new HTTP request.
            HttpClient httpClient = new HttpClient();

            // Set the HTTP authorization header using the access token.
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            // Finish setting up the HTTP request.
            HttpRequestMessage req = new HttpRequestMessage(HttpMethod.Post, url);
            req.Headers.Add("SOAPAction", SOAPAction);
            req.Method = HttpMethod.Post;
            req.Content = new StringContent(xmlSOAP);
            req.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("text/xml; charset=utf-8");

            // Send the request asychronously and wait for the response.
            HttpResponseMessage response;
            response = await httpClient.SendAsync(req);
            var responseBodyAsText = await response.Content.ReadAsStringAsync();

            return responseBodyAsText;
        }
    }
}
//</snippetModernSoapApp2>
