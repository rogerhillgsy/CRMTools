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
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Description;

namespace Microsoft.Crm.Sdk.Samples
{
	using CrmSdk;
	using CrmSdk.Discovery;

	internal static class Program
	{
		#region Constants
		/// <summary>
		/// User Domain
		/// </summary>
		private const string UserDomain = "mydomain";

		/// <summary>
		/// User Name
		/// </summary>
		private const string UserName = "username";

		/// <summary>
		/// Password
		/// </summary>
		private const string UserPassword = "password";

		/// <summary>
		/// Unique Name of the organization
		/// </summary>
		private const string OrganizationUniqueName = "orgname";

		/// <summary>
		/// URL for the Discovery Service
		/// </summary>
		private const string DiscoveryServiceUrl = "http://myserver/XRMServices/2011/Discovery.svc";
		#endregion

		static void Main(string[] args)
		{
			//Generate the credentials
			ClientCredentials credentials = new ClientCredentials();
			credentials.Windows.ClientCredential = new NetworkCredential(UserName, UserPassword, UserDomain);

			//Execute the sample
			string serviceUrl = DiscoverOrganizationUrl(credentials, OrganizationUniqueName, DiscoveryServiceUrl);
			ExecuteWhoAmI(credentials, serviceUrl);
		}

		private static string DiscoverOrganizationUrl(ClientCredentials credentials, string organizationName, string discoveryServiceUrl)
		{
			using (DiscoveryServiceClient client = new DiscoveryServiceClient("CustomBinding_IDiscoveryService", discoveryServiceUrl))
			{
				ApplyCredentials(client, credentials);

				RetrieveOrganizationRequest request = new RetrieveOrganizationRequest()
				{
					UniqueName = organizationName
				};
				RetrieveOrganizationResponse response = (RetrieveOrganizationResponse)client.Execute(request);
				foreach (KeyValuePair<CrmSdk.Discovery.EndpointType, string> endpoint in response.Detail.Endpoints)
				{
                    if (CrmSdk.Discovery.EndpointType.OrganizationService == endpoint.Key)
					{
						Console.WriteLine("Organization Service URL: {0}", endpoint.Value);
						return endpoint.Value;
					}
				}

				throw new InvalidOperationException(string.Format(CultureInfo.InvariantCulture,
					"Organization {0} does not have an OrganizationService endpoint defined.", organizationName));
			}
		}

		private static void ExecuteWhoAmI(ClientCredentials credentials, string serviceUrl)
		{
			using (OrganizationServiceClient client = new OrganizationServiceClient("CustomBinding_IOrganizationService", new EndpointAddress(serviceUrl)))
			{
				ApplyCredentials(client, credentials);

				OrganizationRequest request = new OrganizationRequest();
				request.RequestName = "WhoAmI";

				OrganizationResponse response = (OrganizationResponse)client.Execute(request);

				foreach (KeyValuePair<string, object> result in response.Results)
				{
					if ("UserId" == result.Key)
					{
						Console.WriteLine("User ID: {0}", result.Value);
						break;
					}
				}
			}
		}

		private static void ApplyCredentials<TChannel>(ClientBase<TChannel> client, ClientCredentials credentials)
			where TChannel : class
		{
			client.ClientCredentials.Windows.ClientCredential = credentials.Windows.ClientCredential;
		}
	}
}