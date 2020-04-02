using System;
using System.Collections.Generic;
using System.ServiceModel.Description;
using System.Linq;

// These namespaces are found in the Microsoft.Xrm.Sdk.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;

namespace Microsoft.Crm.Sdk.Samples
{
    public class ConsolelessServerConnection : ServerConnection
    {

        #region Private properties

        private Configuration config = new Configuration();

        #endregion Private properties

        public virtual ServerConnection.Configuration GetServerConfiguration(string server, string orgName, string user, string pw, string domain )
        {
            config.ServerAddress = server;
            if (config.ServerAddress.EndsWith(".dynamics.com"))
            {
                config.EndpointType = AuthenticationProviderType.LiveId;
                config.DiscoveryUri =
                    new Uri(String.Format("https://dev.{0}/XRMServices/2011/Discovery.svc", config.ServerAddress));
                config.DeviceCredentials = GetDeviceCredentials();
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = user;
                credentials.UserName.Password = pw;
                config.Credentials = credentials;
                config.OrganizationUri = GetOrganizationAddress(config.DiscoveryUri, orgName);
            }
            else if (config.ServerAddress.EndsWith(".com"))
            {
                config.EndpointType = AuthenticationProviderType.Federation;
                config.DiscoveryUri =
                    new Uri(String.Format("https://{0}/XRMServices/2011/Discovery.svc", config.ServerAddress));
                ClientCredentials credentials = new ClientCredentials();
                credentials.Windows.ClientCredential = new System.Net.NetworkCredential(user, pw, domain);
                config.Credentials = credentials;
                config.OrganizationUri = GetOrganizationAddress(config.DiscoveryUri, orgName);
            }
            else
            {
                config.EndpointType = AuthenticationProviderType.ActiveDirectory;
                config.DiscoveryUri =
                    new Uri(String.Format("http://{0}/XRMServices/2011/Discovery.svc", config.ServerAddress));
                ClientCredentials credentials = new ClientCredentials();
                credentials.Windows.ClientCredential = new System.Net.NetworkCredential(user, pw, domain);
                config.Credentials = credentials;
                config.OrganizationUri = GetOrganizationAddress(config.DiscoveryUri, orgName);
            }

            if (configurations == null) configurations = new List<Configuration>();
            configurations.Add(config);

            return config;
        }

        protected virtual Uri GetOrganizationAddress(Uri discoveryServiceUri, string orgName)
        {
            using (DiscoveryServiceProxy serviceProxy = new DiscoveryServiceProxy(discoveryServiceUri, null, config.Credentials, config.DeviceCredentials))
            {
                // Obtain organization information from the Discovery service. 
                if (serviceProxy != null)
                {
                    // Obtain information about the organizations that the system user belongs to.
                    OrganizationDetailCollection orgs = DiscoverOrganizations(serviceProxy);

                    OrganizationDetail org = orgs.Where(x => x.UniqueName.ToLower() == orgName.ToLower()).FirstOrDefault();

                    if (org != null)
                    {
                        return new System.Uri(org.Endpoints[EndpointType.OrganizationService]);
                    }
                    else
                    {
                        throw new InvalidOperationException("That OrgName does not exist on that server.");
                    }
                }
                else
                    throw new Exception("An invalid server name was specified.");
            }
        }
    }
}
