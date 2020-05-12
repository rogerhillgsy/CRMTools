using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using System.Configuration;
using System.IO;
using Microsoft.Xrm.Tooling.Connector;
using System.ServiceModel.Description;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Azure.KeyVault;

namespace UpdateOrgsCASLBatch
{
    class Helper : IConfig
    {
        private static string _keyVaultPath;
        private static string _CRMHubPWKey = String.Empty;
        private static Task<string> _CRMHubPWTask;

        string IConfig.KeyVaultPath => KeyVaultPath;

        public static string KeyVaultPath
        {
            get
            {
                if (String.IsNullOrEmpty(_keyVaultPath))
                {
                    _keyVaultPath = ConfigurationManager.AppSettings.Get("KeyVaultPath");
                }

                return _keyVaultPath;
            }
        }

        public string CRMHubPWKey
        {
            get
            {
                _CRMHubPWTask.Wait();
                return _CRMHubPWKey;
            }
        }

        static string password = string.Empty;
        public void connectToCRM()
        {
            _CRMHubPWTask = GetSecretTask("CrmHub-Password", s => _CRMHubPWKey = s);
            _CRMHubPWTask.Wait();
            string environment = ConfigurationManager.AppSettings["Environment"].ToString();
            //CrmServiceClient crmSvc = new CrmServiceClient(ConfigurationManager.ConnectionStrings[environment].ConnectionString);

            //string userId = ConfigurationManager.AppSettings["UserId"].ToString();
            //string pass = ConfigurationManager.AppSettings["Password"].ToString();
            //string URL = ConfigurationManager.AppSettings["URL"].ToString();
            //var connection = CrmConnection.Parse("Url=" + URL + "; Domain=arup.com; Username=" + userId + "; Password=" + pass + ";");
            //OrganizationService service = new OrganizationService(crmSvc);

            string serverUrl = ConfigurationManager.ConnectionStrings[environment].ConnectionString;
            string userName = ConfigurationManager.AppSettings["UserName"].ToString();
            //string password1 = "CIm2$98pRt";
            string domain = ConfigurationManager.AppSettings["Domain"].ToString();
            
            IOrganizationService service = CreateService(serverUrl, userName, password, domain);

            List<Guid> orgsForUpdate = GetOrgsForUpdate(service);
            int count = UpdateFilteredOrgs(orgsForUpdate, service);
            Console.WriteLine(count + " record(s) updated");
        }

        public static IOrganizationService CreateService(string serverUrl, string userId, string password, string domain)
        {
            
            Console.WriteLine("\n\nConnecting to CRM..........\n\n");

            //objDataValidation.CreateLog("Before CRm  Creation");
            IOrganizationService _service;

            ClientCredentials Credentials = new ClientCredentials();
            ClientCredentials devivceCredentials = new ClientCredentials();

            Credentials.UserName.UserName = domain + "\\" + userId;
            //Credentials.UserName.UserName = userId;
            Credentials.UserName.Password = password;
            Uri OrganizationUri = new Uri(serverUrl);
            //Here I am using APAC.
            Uri HomeRealmUri = null;
            //To get device id and password.
            //Online: For online version, we need to call this method to get device id.
            try
            {
                if (!string.IsNullOrEmpty(serverUrl) && serverUrl.Contains("https"))
                {
                    ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                }
                using (OrganizationServiceProxy serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credentials, devivceCredentials))
                {
                    //serviceProxy.ClientCredentials.UserName.UserName = userId; // Your Online username.Eg:username@yourco mpany.onmicrosoft.com";
                    //serviceProxy.ClientCredentials.UserName.Password = password; //Your Online password
                    serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.EndpointBehaviors.Add(new ProxyTypesBehavior());
                    serviceProxy.Timeout = new TimeSpan(0, 120, 0);
                    _service = (IOrganizationService)serviceProxy;
                }
                Console.WriteLine("Connection Established!!!\n\n");
                return _service;
            }
            catch (Exception ex)
            {
                throw new System.Exception("<Error>Problem in creating CRM Service</Error>" + ex.Message);
            }
        }

        public static async Task<string> GetSecretTask(string secretName, Action<string> callback)
        {
            try
            {
                var azureServiceTokenprovider = new AzureServiceTokenProvider();
                var keyVaultClient =
                    new KeyVaultClient(
                        new KeyVaultClient.AuthenticationCallback(azureServiceTokenprovider.KeyVaultTokenCallback));
                //Log($"Accessing Key vault path {KeyVaultPath.TrimEnd("/".ToCharArray())}/secrets/{secretName}");
                var result = String.Empty;

                var getSecretTask = keyVaultClient
                    .GetSecretAsync($"{KeyVaultPath.TrimEnd("/".ToCharArray())}/secrets/{secretName}").ContinueWith(t =>
                    {
                        if (t.IsFaulted)
                        {
                            //Log($"GetSecretTask failed : {t.Exception.Message}");
                            foreach (var exception in t.Exception.InnerExceptions)
                            {
                                //Log($"  {exception.Message}");
                            }
                        }
                        else
                        {
                            result = t.Result.Value;
                            callback?.Invoke(t.Result.Value);
                            password = t.Result.Value;
                            //processRecords();
                            //callback.Invoke()
                        }
                    }
                    );
                await getSecretTask;

                // example: - "https://crmcloudkeys.vault.azure.net/secrets/oracle-test-connection"

                //Log($"Obtained secret \"{secretName}\" from keyvault");
                return result;
            }
            catch (Exception ex)
            {
                //Log($"Failed to get secret ${secretName} message: {ex.Message}");
                throw;
            }
        }
        private List<Guid> GetOrgsForUpdate(IOrganizationService service)
        {
            string fetchXMLCanadaOrgs = @"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='account'>
                    <attribute name='accountid' />
                    <filter type='and'>
                      <condition attribute='statecode' operator='eq' value='0' />
                      <condition attribute='ccrm_countryid' operator='eq' uiname='Canada' uitype='ccrm_country' value='{96B739B0-F03C-E011-9A14-78E7D1644F78}' />
                      <filter type='or'>
                        <condition attribute='arup_impliedconsentexpirydate' operator='olderthan-x-days' value='1' />
                        <condition attribute='arup_impliedconsentexpirydate' operator='yesterday' />
                      </filter>
                      <filter type='or' >
                        <condition attribute='arup_allowcommunication' operator='eq' value='1' />
                        <condition attribute='arup_impliedconsent' operator='eq' value='1' />
                        <condition attribute='arup_lastopportunity' operator='not-null' />
                      </filter>
                    </filter>
              </entity>
            </fetch>";
            List<Guid> FilteredOrgsGuid = new List<Guid>();
            try
            {
                EntityCollection orgsCanada = service.RetrieveMultiple(new FetchExpression(fetchXMLCanadaOrgs));
                Console.WriteLine("\n" + orgsCanada.Entities.Count + " Orgs found for the mentioned criteria");
                List<Organisation> orgsFetchedList = new List<Organisation>();
                foreach (Entity item in orgsCanada.Entities)
                {
                    Entity organisation = new Entity("account");
                    // Create a column set to define which attributes should be retrieved.
                    ColumnSet attributes = new ColumnSet(new string[] { "accountid" });

                    // Retrieve the opportunity.
                    organisation = service.Retrieve(organisation.LogicalName, item.Id, attributes);
                    //Entity rootOrganisation = new Entity("account");

                    Guid rootOrganisation = GetRootForThisOrg(organisation.Id, service);
                    Organisation orgItem = new Organisation();
                    orgItem.AccountId = organisation.Id;
                    orgItem.RootParentId = rootOrganisation;
                    orgsFetchedList.Add(orgItem);

                }
                var groupedOrgs = from org in orgsFetchedList
                                  group org by org.RootParentId into orgx
                                  select new { accountid = orgx.Max(x => x.AccountId) };

               
                foreach (var org in groupedOrgs)
                {
                    FilteredOrgsGuid.Add(org.accountid);
                }
            }
            catch (Exception ex)
            {

                LogExceptions(ex.Message, "N/A", "GetOrgsForUpdate");
            }
            
            
            return FilteredOrgsGuid;

        }
        private Guid GetRootForThisOrg(Guid accountId, IOrganizationService service)
        {
            string fetchXML = "<fetch>" +
                                  "<entity name='account' >" +
                                    "<attribute name='accountid' />" +
                                    "<filter>" +
                                      "<condition attribute='accountid' operator='above' value='" + accountId + "' />" +
                                      "<condition attribute='parentaccountid' operator='null' />" +
                                    "</filter>" +
                                  "</entity>" +
                                "</fetch>";
            try
            {
                EntityCollection account = service.RetrieveMultiple(new FetchExpression(fetchXML));
                if (account != null && account.Entities.Count > 0)
                {
                    return account.Entities[0].Id;
                }
            }
            catch (Exception ex)
            {
                LogExceptions(ex.Message, "N/A", "GetRootForThisOrg");
            }
            
            return accountId;
        }
        private int UpdateFilteredOrgs(List<Guid> guidList, IOrganizationService service)
        {
            int count = 0;
            foreach (var id in guidList)
            {
                try
                {
                    bool expressedConsentAbsent = false, impliedConsentAbsent = false;
                    ColumnSet attributes = new ColumnSet(new string[] { "accountid",  "arup_impliedconsent", "arup_expressedconsent", "arup_allowcommunication","arup_lastopportunity", "arup_impliedconsentexpirydate"});
                    Entity OrgForUpdate = service.Retrieve("account", id, attributes);
                    if (!OrgForUpdate.Attributes.Keys.Contains("arup_impliedconsent"))
                    {
                        OrgForUpdate.Attributes.Add("arup_impliedconsent", null);
                        impliedConsentAbsent = true;
                    }
                    if (!OrgForUpdate.Attributes.Keys.Contains("arup_expressedconsent"))
                    {
                        OrgForUpdate.Attributes.Add("arup_expressedconsent", null);
                        expressedConsentAbsent = true;
                    }
                    if (!OrgForUpdate.Attributes.Keys.Contains("arup_allowcommunication"))
                    {
                        OrgForUpdate.Attributes.Add("arup_allowcommunication", null);
                    }
                    if (impliedConsentAbsent == false)
                    {
                        OrgForUpdate.Attributes["arup_impliedconsent"] = false;
                        if ((Convert.ToBoolean(OrgForUpdate.Attributes["arup_expressedconsent"]) == false) && (expressedConsentAbsent == false))
                        {
                            OrgForUpdate.Attributes["arup_allowcommunication"] = false;
                        }
                    }
                    if (OrgForUpdate.Attributes.Keys.Contains("arup_lastopportunity"))
                    {
                        OrgForUpdate.Attributes["arup_lastopportunity"] = null;
                    }
                    if (OrgForUpdate.Attributes.Keys.Contains("arup_impliedconsentexpirydate"))
                    {
                        OrgForUpdate.Attributes["arup_impliedconsentexpirydate"] = null;
                    }
                    service.Update(OrgForUpdate);
                    count++;
                }
                catch (Exception ex)
                {
                    LogExceptions(ex.Message, id.ToString(), "UpdateFilteredOrgs");
                }
            }
            return count;
        }
        private void LogExceptions(string message, string guid,string methodName)
        {
            string fileName = DateTime.Now.ToString("yyyyMMdd")+".csv";
            using (StreamWriter writer = new StreamWriter(fileName,true))
            {
                writer.WriteLine(guid+","+methodName+","+message);
            }
        }
        
    }
}
