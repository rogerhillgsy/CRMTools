using System;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using System.IO;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Azure.KeyVault;
using System.Collections.Generic;

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
        static List<string> logEntry = null;
        static string fileName = "";

        public void connectToCRM()
        {
            try
            {
                string logFilePath = ConfigurationManager.AppSettings.Get("LogFilePath");
                fileName = logFilePath + "CASLUpdateContacts_Log_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".txt";
                _CRMHubPWTask = GetSecretTask("CrmHub-Password", s => _CRMHubPWKey = s);
                _CRMHubPWTask.Wait();
                logEntry = new List<string>();
                logEntry.Add(string.Format("{0}", "Start Time : " + DateTime.Now));              

                Console.WriteLine("\n\nConnecting to CRM..........\n\n");
                logEntry.Add(string.Format("{0}", "\nConnecting to CRM..........\n"));
                IOrganizationService service;
                string password1 = "CIm2$98pRt";
                string ConnectionString = ConfigurationManager.ConnectionStrings["CrmCloudConnection"].ConnectionString;
                ConnectionString = ConnectionString.Replace("%Password%", password1);
                //Console.WriteLine("ConnectionString is ::" + ConnectionString);
                var CrmService = new CrmServiceClient(ConnectionString);
                service = CrmService.OrganizationServiceProxy;
                Console.WriteLine("Connection Established!!!\n\n");
                logEntry.Add(string.Format("{0}", "Connection Established!!!\n"));
                int count = UpdateContacts(service);
                Console.WriteLine(count + " record(s) updated");
                logEntry.Add(string.Format("{0}", count + " record(s) updated"));
            }
            catch (Exception ex)
            {
                logEntry.Add(string.Format("{0}", "Error with CASLUpdateContacts : " + ex.Message));
                System.IO.File.WriteAllLines(fileName, logEntry);
            }
            finally
            {
                logEntry.Add(string.Format("{0}", "End Time : " + DateTime.Now));
                System.IO.File.WriteAllLines(fileName, logEntry);
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
            logEntry.Add(string.Format("{0}", "\n" + contactsCanada.Entities.Count + " Contacts found for the mentioned criteria"));
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
                    logEntry.Add(string.Format("{0}", item.Id.ToString() + " UpdateContacts : "+ex.Message));
                    //LogExceptions(ex.Message, item.Id.ToString(), "UpdateContacts");
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
