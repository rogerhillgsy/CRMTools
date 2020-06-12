using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.IO;
using System.Data;
using Microsoft.Xrm.Sdk.Query;
using System.Configuration;
using System.Collections.Specialized;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using System.Text.RegularExpressions;

namespace OrganizationClientSectorClientTypeMapping
{
    public class Program : IConfig
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
        static List<string> linesInFailedFile = null;
        static string fileName = string.Empty;
        static string password = string.Empty;
        static string entityname = string.Empty;
        static string fromAttribute = string.Empty;
        static string toAttribute = string.Empty;
        static int totalRejectedRecordCount = 0;
        static int totalSuccessRecordCount = 0;
        public static void Main(string[] args)
        {
            try
            {
                _CRMHubPWTask = GetSecretTask("CrmHub-Password", s => _CRMHubPWKey = s);
                _CRMHubPWTask.Wait();
                processRecords();
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Error occured : ", ex.InnerException.Message);
                Console.WriteLine("Error occured : {0} ", ex.Message);
            }

        }

        public static void processRecords()
        {
            try
            {
                string serverUrl = ConfigurationManager.AppSettings["serverUrl"].ToString();
                string userName = ConfigurationManager.AppSettings["UserName"].ToString();
                //string password = ConfigurationManager.AppSettings["Password"].ToString();
                string password1 = "CIm2$98pRt";
                string domain = ConfigurationManager.AppSettings["Domain"].ToString();
                entityname = ConfigurationManager.AppSettings["Entity"].ToString();
                fromAttribute = ConfigurationManager.AppSettings["FromAttribute"].ToString();
                toAttribute = ConfigurationManager.AppSettings["ToAttribute"].ToString();
                fileName = entityname + "_FailedRecordsLogFile_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".csv";
                linesInFailedFile = new List<string>();
                linesInFailedFile.Add("Entity,RecordId,Error Description, OptionSetValues");
                linesInFailedFile.Add(string.Format("{0}", "Start Time : " + DateTime.Now));
                Console.WriteLine(entityname + " Entity processing : Start time:" + DateTime.Now);
                IOrganizationService service = CreateService(serverUrl, userName, password1, domain);
                //IOrganizationService service = CreateService1("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc", "crm.hub@arup.com", "CIm2$98pRt", "arup");
                UpdateEntityOptionSet(service, fromAttribute);

            }
            catch (Exception ex)
            {
                //Console.WriteLine("Error occured : ", ex.InnerException.Message);
                Console.WriteLine("Bid Review : Error occured : {0} ", ex.Message);
                linesInFailedFile.Add(string.Format("{0}", ex.Message));
            }
            finally
            {
                System.IO.File.WriteAllLines(fileName, linesInFailedFile);
            }
        }

        #region CRM CONNECTION


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

        #endregion

        #region code with skipping processed records
        public static void UpdateEntityOptionSet(IOrganizationService service, string fromAttribute)
        {
            try
            {
                QueryExpression query = new QueryExpression(entityname);
                string entityprimaryid = entityname + "id";

                query.ColumnSet.AddColumns("accountid");
                query.Criteria = new FilterExpression();

                FilterExpression filter = new FilterExpression(LogicalOperator.And);
                FilterExpression filter1 = new FilterExpression(LogicalOperator.Or);
                filter1.Conditions.Add(new ConditionExpression(fromAttribute, ConditionOperator.Like, "%42%"));
                filter1.Conditions.Add(new ConditionExpression(fromAttribute, ConditionOperator.Like, "%43%"));
                filter1.Conditions.Add(new ConditionExpression(fromAttribute, ConditionOperator.Like, "%44%"));
                filter.AddFilter(filter1);
                query.Criteria = filter;

                query.PageInfo = new PagingInfo();
                query.PageInfo.Count = 1;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;
                //
                //QueryExpressionToFetchXmlRequest req = new QueryExpressionToFetchXmlRequest();
                //req.Query = query;
                //QueryExpressionToFetchXmlResponse resp = (QueryExpressionToFetchXmlResponse)service.Execute(req);

                //string myfetch = resp.FetchXml;
                //
                EntityCollection entityCollection = service.RetrieveMultiple(query);
                
                int rowCount = 0;
                rowCount = entityCollection.Entities.Count;
                for (int i = 0; i < entityCollection.Entities.Count; i++)
                {
                    Entity entity = new Entity(entityname);
                    entity[toAttribute] = new OptionSetValue(100000005);
                    entity.Id = entityCollection.Entities[i].GetAttributeValue<Guid>(entityprimaryid);
                    service.Update(entity);
                    totalSuccessRecordCount++;
                }
                if (rowCount > 0)
                {
                    System.IO.File.WriteAllLines(fileName, linesInFailedFile);
                    int totalrecordsprocessed1 = totalSuccessRecordCount + totalRejectedRecordCount;
                    Console.WriteLine(totalrecordsprocessed1 + " " + entityname + "  Records processed at : " + DateTime.Now);
                }
                //rowCount = 0;
                do
                {                   
                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                    entityCollection = service.RetrieveMultiple(query);
                    rowCount = entityCollection.Entities.Count;
                    for (int i = 0; i < entityCollection.Entities.Count; i++)
                    {
                        Entity entity = new Entity(entityname);
                        entity["ccrm_clienttype"] = new OptionSetValue(100000005);
                        entity.Id = entityCollection.Entities[i].GetAttributeValue<Guid>(entityprimaryid);
                        service.Update(entity);
                        totalSuccessRecordCount++;
                    }

                    if (entityCollection.Entities.Count > 0)
                        Console.WriteLine(totalSuccessRecordCount + " " + entityname + "  Records processed at : " + DateTime.Now);
                    
                    System.IO.File.WriteAllLines(fileName, linesInFailedFile);                    
                }
                while (entityCollection.MoreRecords);

                int totalrecordsprocessed = totalSuccessRecordCount + totalRejectedRecordCount;
                linesInFailedFile.Add(string.Format("{0}", "Total " + entityname + " records processed count:" + totalrecordsprocessed));
                linesInFailedFile.Add(string.Format("{0}", "Success Count : " + totalSuccessRecordCount + " Failed Count : " + totalRejectedRecordCount));
                linesInFailedFile.Add(string.Format("{0}", "End Time : " + DateTime.Now));
                Console.WriteLine(string.Format("{0}", "Total " + entityname + " records processed count:" + totalrecordsprocessed));
                Console.WriteLine(string.Format("{0}", "Success Count : " + totalSuccessRecordCount + " Failed Count : " + totalRejectedRecordCount));
                Console.WriteLine(string.Format("{0}", "End Time : " + DateTime.Now));
                System.IO.File.WriteAllLines(fileName, linesInFailedFile);
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                totalRejectedRecordCount++;
                //Console.WriteLine("Error occured : ", ex.InnerException.Message);
                Console.WriteLine(entityname + " Entity processing Error occured : {0} ", ex.Message);
                linesInFailedFile.Add(string.Format("{0}", ex.Message));
                System.IO.File.WriteAllLines(fileName, linesInFailedFile);
            }
        }
        #endregion



        //public static Dictionary<Nullable<Int32>, string> RetriveOptionSetLabels(IOrganizationService service, string entityLogicalName, string optionSetLogicalName)
        //{
        //    Dictionary<Nullable<Int32>, string> dic = new Dictionary<int?, string>();
        //    string EntityLogicalName = entityLogicalName;
        //    string FieldLogicalName = optionSetLogicalName;

        //    string FetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >";
        //    FetchXml = FetchXml + "<entity name='stringmap' >";
        //    FetchXml = FetchXml + "<attribute name='attributevalue' />";
        //    FetchXml = FetchXml + "<attribute name='value' />";
        //    FetchXml = FetchXml + "<filter type='and' >";
        //    FetchXml = FetchXml + "<condition attribute='objecttypecode' operator='eq' value='" + objectTypeCode + "' />";
        //    FetchXml = FetchXml + "<condition attribute='attributename' operator='eq' value='" + FieldLogicalName + "' />";
        //    FetchXml = FetchXml + "</filter></entity></fetch>";

        //    FetchExpression FetchXmlQuery = new FetchExpression(FetchXml);

        //    EntityCollection FetchXmlResult = service.RetrieveMultiple(FetchXmlQuery);

        //    if (FetchXmlResult.Entities.Count > 0)
        //    {
        //        foreach (Entity Stringmap in FetchXmlResult.Entities)
        //        {
        //            string OptionValue = Stringmap.Attributes.Contains("value") ? (string)Stringmap.Attributes["value"] : string.Empty;
        //            Int32 OptionLabel = Stringmap.Attributes.Contains("attributevalue") ? (Int32)Stringmap.Attributes["attributevalue"] : 0;
        //            dic.Add(OptionLabel, OptionValue);
        //        }
        //    }
        //    return dic;
        //}



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
