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
using ArupMultiSelect;

namespace CloseOpportunityReasons
{
    class Program : IConfig
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
        static List<string> linesInFailedFile = null;
        static string fileName = "FailedRecordsFile" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".csv";
        static int StartPageNumber = 0;
        static int EndPageNumber = 0;
        static int RecordCountPerPage = 0;
        static void Main(string[] args)
        {
            try
            {
                _CRMHubPWTask = GetSecretTask("CrmHub-Password", s => _CRMHubPWKey = s);
                _CRMHubPWTask.Wait();
                Console.WriteLine("Close Opportunity Reasons records Processing Statred. Start time:" + DateTime.Now);
                linesInFailedFile = new List<string>();
                linesInFailedFile.Add("Entity,RecordId,Error Description, OptionSetValues");
                linesInFailedFile.Add(string.Format("{0},{1},{2},{3}", "Start Time : " + DateTime.Now, "", "", ""));
                string serverUrl = ConfigurationManager.AppSettings["serverUrl"].ToString();
                string userName = ConfigurationManager.AppSettings["UserName"].ToString();
                //string password = ConfigurationManager.AppSettings["Password"].ToString();CIm2$98pRt
                password = "CIm2$98pRt";
                string domain = ConfigurationManager.AppSettings["Domain"].ToString();
                StartPageNumber = Convert.ToInt32(ConfigurationManager.AppSettings["StartPageNumber"]);
                EndPageNumber = Convert.ToInt32(ConfigurationManager.AppSettings["EndPageNumber"]);
                RecordCountPerPage = Convert.ToInt32(ConfigurationManager.AppSettings["RecordCountPerPage"]);
                IOrganizationService service = CreateService(serverUrl, userName, password, domain);
                //IOrganizationService service = CreateService("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc", "crm.hub@arup.com", "CIm2$98pRt", "arup");
                UpdateCloseopportunityreason(service);

                linesInFailedFile.Add(string.Format("{0},{1},{2},{3}", "End Time : " + DateTime.Now, "", "", ""));
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Error occured : ", ex.InnerException.Message);
                Console.WriteLine("Error occured : {0} ", ex.Message);
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
                    serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());
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



        #region Update Opportunity
        public static void UpdateCloseopportunityreason(IOrganizationService service)
        {
            QueryExpression test = new QueryExpression();

            QueryExpression query = new QueryExpression("arup_closeopportunityreason");
            //query.ColumnSet = new ColumnSet("ccrm_businessinterestpicklistname", "ccrm_businessinterestpicklistvalue", "arup_businessinterest");
            query.ColumnSet.AddColumns("arup_lostopportunityreasonvalues", "arup_wonopportunityreasonvalues");
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.Or;
            query.Criteria.AddCondition("arup_lostopportunityreasonvalues", ConditionOperator.NotNull);
            query.Criteria.AddCondition("arup_wonopportunityreasonvalues", ConditionOperator.NotNull);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = RecordCountPerPage;
            query.PageInfo.PageNumber = StartPageNumber;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            EntityCollection final = new EntityCollection();
            foreach (Entity i in entityCollection.Entities)
            {
                final.Entities.Add(i);
                UpdateLeadMultiSelect(service,
                    i.GetAttributeValue<Guid>("arup_closeopportunityreasonid"),
                    i.GetAttributeValue<string>("arup_lostopportunityreasonvalues"),
                    i.GetAttributeValue<string>("arup_wonopportunityreasonvalues"));
            }
            do
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                entityCollection = service.RetrieveMultiple(query);
                foreach (Entity i in entityCollection.Entities)
                {
                    final.Entities.Add(i);
                    UpdateLeadMultiSelect(service,
                    i.GetAttributeValue<Guid>("arup_closeopportunityreasonid"),
                    i.GetAttributeValue<string>("arup_lostopportunityreasonvalues"),
                    i.GetAttributeValue<string>("arup_wonopportunityreasonvalues"));
                }

                Console.WriteLine(query.PageInfo.PageNumber * RecordCountPerPage + " Close Opportunity Reasons Records processed at : " + DateTime.Now);
                System.IO.File.WriteAllLines(fileName, linesInFailedFile);
                if (query.PageInfo.PageNumber == EndPageNumber)
                    break;
            }
            while (entityCollection.MoreRecords);
            Console.WriteLine("Total Close Opportunity Reasons record count:" + RecordCountPerPage * EndPageNumber);
            Console.WriteLine("Close Opportunity Reasons records Processing Completed. End time:" + DateTime.Now);
            Console.ReadKey();
        }

        //ccrm_othernetworksval", "ccrm_servicesvalue", "ccrm_theworksvalue", "ccrm_disciplinesvalue", "ccrm_projectsectorvalue"
        public static void UpdateLeadMultiSelect(IOrganizationService service, Guid arup_closeopportunityreasonid, string arup_lostopportunityreasonvalues, string arup_wonopportunityreasonvalues)
        {
            try
            {
                Entity opportunity = new Entity("arup_closeopportunityreason");
                if (arup_lostopportunityreasonvalues != string.Empty && arup_lostopportunityreasonvalues != null)
                {
                    Dictionary<Nullable<int>, string> opset = RetriveOptionSetLabels(service, "arup_closeopportunityreason", "arup_lostopportunityreasonslist");
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_lostopportunityreasonvalues.Split(',');
                    foreach (var item in arr)
                    {
                        if (item != null && item.Trim() != string.Empty && item.Trim() != "" && item.All(char.IsDigit) == true)
                        {
                            if (opset.ContainsKey(Convert.ToInt32(item)))
                            {
                                collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                            }
                        }
                    }

                    opportunity["arup_lostreasons"] = collectionOptionSetValues;
                }
                if (arup_wonopportunityreasonvalues != string.Empty && arup_wonopportunityreasonvalues != null)
                {
                    Dictionary<Nullable<int>, string> opset = RetriveOptionSetLabels(service, "arup_closeopportunityreason", "arup_wonopportunityreasonslist");
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_wonopportunityreasonvalues.Split(',');
                    foreach (var item in arr)
                    {
                        if (item != null && item.Trim() != string.Empty && item.Trim() != "" && item.All(char.IsDigit) == true)
                        {
                            if (opset.ContainsKey(Convert.ToInt32(item)))
                            {
                                collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                            }
                        }
                    }

                    opportunity["arup_wonreasons"] = collectionOptionSetValues;
                }

                opportunity.Id = arup_closeopportunityreasonid;
                service.Update(opportunity);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error : " + e.Message);
                //linesInFailedFile.Add("RecordId,ToOptionset,Values,Message");
                string optionSetValues = "arup_lostreasons : " + arup_lostopportunityreasonvalues + " | arup_wonreasons : " + arup_wonopportunityreasonvalues;

                linesInFailedFile.Add(string.Format("{0},{1},{2},{3}", "arup_closeopportunityreason", arup_closeopportunityreasonid, e.Message, optionSetValues));
            }
        }

        public static Dictionary<Nullable<Int32>, string> RetriveOptionSetLabels(IOrganizationService service, string entityLogicalName, string optionSetLogicalName)
        {

            //var attributeRequest = new RetrieveAttributeRequest
            //{
            //    EntityLogicalName = entityLogicalName,
            //    LogicalName = optionSetLogicalName,
            //    RetrieveAsIfPublished = true
            //};

            //var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
            //var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

            //var optionList = (from o in attributeMetadata.OptionSet.Options
            //                  select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();


            Dictionary<Nullable<Int32>, string> dic = new Dictionary<int?, string>();
            string EntityLogicalName = entityLogicalName;
            string FieldLogicalName = optionSetLogicalName;

            string FetchXml = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false' >";
            FetchXml = FetchXml + "<entity name='stringmap' >";
            FetchXml = FetchXml + "<attribute name='attributevalue' />";
            FetchXml = FetchXml + "<attribute name='value' />";
            FetchXml = FetchXml + "<filter type='and' >";
            FetchXml = FetchXml + "<condition attribute='objecttypecode' operator='eq' value='10142' />";
            FetchXml = FetchXml + "<condition attribute='attributename' operator='eq' value='" + FieldLogicalName + "' />";
            FetchXml = FetchXml + "</filter></entity></fetch>";

            FetchExpression FetchXmlQuery = new FetchExpression(FetchXml);

            EntityCollection FetchXmlResult = service.RetrieveMultiple(FetchXmlQuery);

            if (FetchXmlResult.Entities.Count > 0)
            {
                foreach (Entity Stringmap in FetchXmlResult.Entities)
                {
                    string OptionValue = Stringmap.Attributes.Contains("value") ? (string)Stringmap.Attributes["value"] : string.Empty;
                    Int32 OptionLabel = Stringmap.Attributes.Contains("attributevalue") ? (Int32)Stringmap.Attributes["attributevalue"] : 0;
                    dic.Add(OptionLabel, OptionValue);
                }
            }
            return dic;
        }
        #endregion

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
