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
using System.Configuration;
using System.IO;
using Microsoft.Xrm.Tooling.Connector;
using System.ServiceModel.Description;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Azure.KeyVault;

namespace CASLUpdateContacts
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

            string serverUrl = ConfigurationManager.ConnectionStrings[environment].ConnectionString;
            string userName = ConfigurationManager.AppSettings["UserName"].ToString();
            string password1 = "CIm2$98pRt";
            string domain = ConfigurationManager.AppSettings["Domain"].ToString();

            IOrganizationService service = CreateService(serverUrl, userName, password1, domain);
            //OrganizationService service = new OrganizationService(crmSvc);
            int count = UpdateContacts(service);
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

        private int UpdateContacts(IOrganizationService service)
        {
            string fetchXMLCanadaContacts =
                @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='contact'>
                    <attribute name='contactid' />
                    <attribute name='arup_expirydate' />
                    <attribute name='arup_organisationconsent' />
                    <attribute name='arup_otherimpliedconsent' />
                    <attribute name='arup_expressedconsent' />
                    <attribute name='arup_allowcommunication' />
                    <order attribute='arup_expirydate' descending='false' />
                    <filter type='and'>
                      <condition attribute='statecode' operator='eq' value='0' />
                      <condition attribute='ccrm_countryid' operator='eq' uiname='Canada' uitype='ccrm_country' value='{96B739B0-F03C-E011-9A14-78E7D1644F78}' />
                      <filter type='or'>
                        <condition attribute='arup_expirydate' operator='olderthan-x-days' value='1' />
                        <condition attribute='arup_expirydate' operator='yesterday' />
                      </filter>
                        <condition attribute='arup_otherimpliedconsent' operator='ne' value='770000005' />
                    </filter>
                  </entity>
            </fetch>";
            EntityCollection contactsCanada = service.RetrieveMultiple(new FetchExpression(fetchXMLCanadaContacts));
            Console.WriteLine("\n" + contactsCanada.Entities.Count + " Contacts found for the mentioned criteria");
            int count = 0;
            foreach (Entity item in contactsCanada.Entities)
            {
                try
                {

                    bool arup_expirydateAbsent = false, arup_organisationconsentAbsent = false, arup_otherimpliedconsentAbsent = false, arup_expressedconsentAbsent = false, arup_allowcommunicationAbsent = false;
                    Entity contact = new Entity("contact");
                    ColumnSet attributes = new ColumnSet(new string[] { "contactid", "arup_expirydate", "arup_organisationconsent", "arup_otherimpliedconsent", "arup_expressedconsent", "arup_allowcommunication" });
                    contact = service.Retrieve("contact", item.Id, attributes);
                    OptionSetValue impliedConsentNA = new OptionSetValue(770000005);
                    if (!contact.Attributes.Keys.Contains("arup_expirydate"))
                    {
                        //contact.Attributes.Add("arup_expirydate", null);
                        arup_expirydateAbsent = true;
                    }
                    if (!contact.Attributes.Keys.Contains("arup_organisationconsent"))
                    {
                        contact.Attributes.Add("arup_organisationconsent", null);
                        arup_organisationconsentAbsent = true;
                    }
                    if (!contact.Attributes.Keys.Contains("arup_otherimpliedconsent"))
                    {
                        contact.Attributes.Add("arup_otherimpliedconsent", null);
                        arup_otherimpliedconsentAbsent = true;
                    }
                    if (!contact.Attributes.Keys.Contains("arup_expressedconsent"))
                    {
                        contact.Attributes.Add("arup_expressedconsent", null);
                        arup_expressedconsentAbsent = true;
                    }
                    if (!contact.Attributes.Keys.Contains("arup_allowcommunication"))
                    {
                        contact.Attributes.Add("arup_allowcommunication", null);
                        arup_allowcommunicationAbsent = true;
                    }
                    if (arup_expirydateAbsent == false)
                    {
                        contact["arup_otherimpliedconsent"] = impliedConsentNA;
                    }
                    //if ((Convert.ToBoolean(contact["arup_organisationconsent"]) == true && arup_organisationconsentAbsent == false) || (Convert.ToBoolean(contact["arup_expressedconsent"]) == true && arup_expressedconsentAbsent == false) || (((Microsoft.Xrm.Sdk.OptionSetValue)(contact["arup_otherimpliedconsent"])).Value != 770000005 && arup_otherimpliedconsentAbsent == false))
                    //{
                    //    contact["arup_allowcommunication"] = true;
                    //}
                    //else
                    //{
                    //    contact["arup_allowcommunication"] = false;
                    //}
                    service.Update(contact);
                }
                catch (Exception ex)
                {

                    LogExceptions(ex.Message, item.Id.ToString(), "UpdateContacts");
                }
                count++;

                
            }
            return count;
        }
        private void LogExceptions(string message, string guid, string methodName)
        {
            string fileName = DateTime.Now.ToString("yyyyMMdd") + ".csv";
            using (StreamWriter writer = new StreamWriter(fileName, true))
            {
                writer.WriteLine(guid + "," + methodName + "," + message);
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
    }
}
