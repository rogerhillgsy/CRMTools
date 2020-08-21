using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using System.IO;
using Microsoft.Xrm.Tooling.Connector;
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

        static List<string> logEntry = null;
        static string fileName = "";

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
            try
            {
                fileName = "UpdateOrgsCASLBatch_Log_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".txt";
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
                var CrmService = new CrmServiceClient(ConnectionString);
                service = CrmService.OrganizationServiceProxy;
                Console.WriteLine("Connection Established!!!\n");
                logEntry.Add(string.Format("{0}", "Connection Established!!!\n"));
                List<Guid> orgsForUpdate = GetOrgsForUpdate(service);
                int count = UpdateFilteredOrgs(orgsForUpdate, service);
                Console.WriteLine(count + " record(s) updated");
                logEntry.Add(string.Format("{0}", count + " record(s) updated"));

                //Console.ReadLine();
            }
            catch (Exception ex)
            {
                logEntry.Add(string.Format("{0}", "Error with UpdateOrgsCASLBatch : " + ex.Message));
                System.IO.File.WriteAllLines(fileName, logEntry);
            }
            finally
            {
                logEntry.Add(string.Format("{0}", "End Time : " + DateTime.Now));
                System.IO.File.WriteAllLines(fileName, logEntry);
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
                logEntry.Add(string.Format("{0}", "\n" + orgsCanada.Entities.Count + " Orgs found for the mentioned criteria"));
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
                logEntry.Add(string.Format("{0}", "GetOrgsForUpdate : Exception : " + ex.Message));
                //LogExceptions(ex.Message, "N/A", "GetOrgsForUpdate");
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
                logEntry.Add(string.Format("{0}", "GetRootForThisOrg : Exception : " + ex.Message));
                //LogExceptions(ex.Message, "N/A", "GetRootForThisOrg");
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
                    logEntry.Add(string.Format("{0}", "UpdateFilteredOrgs : Exception : " + ex.Message));
                    //LogExceptions(ex.Message, id.ToString(), "UpdateFilteredOrgs");
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
