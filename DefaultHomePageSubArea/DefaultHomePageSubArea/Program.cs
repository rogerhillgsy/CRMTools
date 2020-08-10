#region Includes
using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Threading.Tasks;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
#endregion

#region namespace  DefaultHomePageSubAreaUpdate
namespace DefaultHomePageSubAreaUpdate
{
    #region class 
    class Program :IConfig
    {
        #region Variables and Properties
        private static IOrganizationService _orgService;

        private static string _keyVaultPath;
        private static string _CRMHubPWKey = String.Empty;
        private static Task<string> _CRMHubPWTask;
        string IConfig.KeyVaultPath => KeyVaultPath;
        static string password = string.Empty;

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
        #endregion

        #region Functions
        /// <summary>
        /// Main function
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            _CRMHubPWTask = GetSecretTask("CrmHub-Password", s => _CRMHubPWKey = s);
            _CRMHubPWTask.Wait();

            string serverUrl = ConfigurationManager.AppSettings["serverUrl"];
            string userId = ConfigurationManager.AppSettings["usersId"];
            string domain = ConfigurationManager.AppSettings["domain"];
            _orgService = CreateService(serverUrl, userId, password, domain);

            // Create a writer and open the file:
            StreamWriter log;
            string logFilename = ConfigurationManager.AppSettings["LogFileName"];

            if (!File.Exists(logFilename))
            {
                log = new StreamWriter(logFilename);
            }
            else
            {
                log = File.AppendText(logFilename);
            }

            Console.WriteLine("Processing of records started at  " + DateTime.Now);
            log.WriteLine("Processing of records started at  " + DateTime.Now);
            log.WriteLine();

            GetAllUserSettingData(_orgService, log);

            Console.WriteLine("Processing of records completed at  " + DateTime.Now);
            log.WriteLine("Processing of records completed at  " + DateTime.Now);
            // Close the stream:
            log.Close();
        }

        /// <summary>
        /// This function is used to retrieve details like Homepagearea and homepagesubarea of all the user where employmentstatus= employee
        /// </summary>
        /// <param name="service">organisationservice</param>
        /// <param name="log"></param>
        static void GetAllUserSettingData(IOrganizationService service, StreamWriter log)
        {
            string defaultHomePageArea = ConfigurationManager.AppSettings["DefaultHomePageArea"];
            string defaultHomePageSubArea = ConfigurationManager.AppSettings["DefaultHomePageSubArea"];

            QueryExpression queryUser = new QueryExpression("systemuser");
            queryUser.ColumnSet = new ColumnSet("systemuserid", "arup_employmentstatus", "arup_defaultdataupdated");
            queryUser.Criteria.AddCondition(new ConditionExpression("arup_employmentstatus", ConditionOperator.Equal, 770000000));
            queryUser.Criteria.AddCondition(new ConditionExpression("arup_defaultdataupdated", ConditionOperator.NotEqual, true));

            queryUser.Criteria.AddCondition(new ConditionExpression("internalemailaddress", ConditionOperator.NotEqual, "Powerqueryonline@onmicrosoft.com"));
            queryUser.Criteria.AddCondition(new ConditionExpression("internalemailaddress", ConditionOperator.NotEqual, "Dynamics365Athena@onmicrosoft.com"));
            queryUser.Criteria.AddCondition(new ConditionExpression("internalemailaddress", ConditionOperator.NotEqual, "Dynamics365Athena2@onmicrosoft.com"));
            queryUser.Criteria.AddCondition(new ConditionExpression("internalemailaddress", ConditionOperator.NotEqual, "Dynamics365EnterpriseSales@onmicrosoft.com"));
            queryUser.Criteria.AddCondition(new ConditionExpression("isdisabled", ConditionOperator.Equal, false)); //enabled user

            LinkEntity lnkUserSetting = new LinkEntity("systemuser", "usersettings", "systemuserid", "systemuserid", JoinOperator.Inner);
            lnkUserSetting.Columns = new ColumnSet("systemuserid", "homepagearea", "homepagesubarea");
            lnkUserSetting.EntityAlias = "usersettings";
            queryUser.LinkEntities.Add(lnkUserSetting);
            queryUser.PageInfo.PagingCookie = null;
            queryUser.PageInfo.Count = Convert.ToInt16(ConfigurationManager.AppSettings["BatchSize"]);

            EntityCollection collUserSettings = new EntityCollection();
            int totalCount = collUserSettings.Entities.Count;
            do
            {

                queryUser.PageInfo.PageNumber += 1;
                queryUser.PageInfo.PagingCookie = collUserSettings.PagingCookie;
                collUserSettings = service.RetrieveMultiple(queryUser);

                Log("-------------------------------", log);
                Log("Retrieved Record for page number " + queryUser.PageInfo.PageNumber, log);
                Log("Total number of Records retrieved " + collUserSettings.Entities.Count, log);

                if (collUserSettings.Entities.Count > 0)
                {
                    UpdateUserRecord(collUserSettings, service, log, queryUser.PageInfo.PageNumber, defaultHomePageArea, defaultHomePageSubArea);
                }

                totalCount = totalCount + collUserSettings.Entities.Count;

            } while (collUserSettings.MoreRecords);

            Log("Total number of Records " + totalCount, log);
        }

        /// <summary>
        /// This function update the defaultpane and defaulttab field of user entity with homepagearea and homepagesubarea respectively. Then update the homepagearea and homepagesubarea with default value "HLP" and "whats new"
        /// </summary>
        /// <param name="collUserSettings">collection of user data</param>
        /// <param name="service">organisationservice</param>
        /// <param name="log"></param>
        /// <param name="pageNumber">pageNumber</param>
        static void UpdateUserRecord(EntityCollection collUserSettings, IOrganizationService service, StreamWriter log, int pageNumber, string defaultHomePageArea, string defaultHomePageSubArea)
        {
         

            ExecuteTransactionRequest transactionRequest = new ExecuteTransactionRequest()
            {
                Requests = new OrganizationRequestCollection()
            };
            try
            {
                Console.WriteLine("Below Set Of {0} Records will be processed for Page {1}", collUserSettings.Entities.Count, pageNumber);
                Log(String.Format("Below Set Of {0} Records will be processed for Page {1}", collUserSettings.Entities.Count, pageNumber), log);

                foreach (Entity userSetting in collUserSettings.Entities)
                {
                    string homePageArea = string.Empty;
                    string homePageSubArea = string.Empty;
                    Guid userId = Guid.Empty;

                    if (userSetting.Attributes.Contains("usersettings.homepagearea"))
                    {
                      
                        homePageArea = ((AliasedValue)userSetting["usersettings.homepagearea"]).Value.ToString();
                        if (homePageArea == "HLP") // This is needed to fix the recursion issue with old record. In Nov release, this block can be removed as there is no 'HLP' homepagearea in Cloud
                        {
                            homePageArea = "<Default>";
                        }
                    }
                    if (userSetting.Attributes.Contains("usersettings.homepagesubarea"))
                    {
                        homePageSubArea = ((AliasedValue)userSetting["usersettings.homepagesubarea"]).Value.ToString();
                        if (homePageSubArea == "Whats_New")
                        {
                            homePageSubArea = "";
                        }
                    }
                    if (userSetting.Attributes.Contains("systemuserid"))
                    {
                        userId = (Guid)userSetting["systemuserid"];
                        Console.WriteLine("UserId :" + userId, log);
                        Log("UserId :" + userId, log);
                    }

                    var retrievedUser = new Entity("systemuser", userId);

                    var user = new Entity("systemuser");
                    user.Id = retrievedUser.Id;
                    user["arup_defaultpane"] = homePageArea;
                    user["arup_defaulttab"] = homePageSubArea;
                    user["arup_defaultdataupdated"] = true;

                    var updateRequestUser = new UpdateRequest()
                    {
                        Target = user
                    };



                    var usersetting = new Entity("usersettings");
                    usersetting.Id = retrievedUser.Id;
                    usersetting["homepagearea"] = defaultHomePageArea;
                    usersetting["homepagesubarea"] = defaultHomePageSubArea;

                    var updateRequestUserSetting = new UpdateRequest()
                    {
                        Target = usersetting
                    };

                    transactionRequest.Requests.Add(updateRequestUser);
                    transactionRequest.Requests.Add(updateRequestUserSetting);
                }

                Console.WriteLine("Before Calling Execute for transactionRequest for page {0}", pageNumber);
                Log(String.Format("Before Calling Execute for transactionRequest for page {0}", pageNumber), log);

               

                 ExecuteTransactionResponse transactResponse = (ExecuteTransactionResponse)service.Execute(transactionRequest);


                Console.WriteLine("After Calling Execute for transactionRequest for page {0}", pageNumber);
                Log(String.Format("After Calling Execute for transactionRequest for page {0}", pageNumber), log);
                Log("--------------------------------", log);

              

            }
            catch (FaultException<OrganizationServiceFault> fault)
            {
                // Check if the maximum batch size has been exceeded. The maximum batch size is only included in the fault if it
                // the input request collection count exceeds the maximum batch size.
                if (fault.Detail.ErrorDetails.Contains("MaxBatchSize"))
                {
                    int maxBatchSize = Convert.ToInt32(fault.Detail.ErrorDetails["MaxBatchSize"]);
                    if (maxBatchSize < transactionRequest.Requests.Count)
                    {
                        // Here you could reduce the size of your request collection and re-submit the ExecuteMultiple request.
                        // For this sample, that only issues a few requests per batch, we will just print out some info. However,
                        // this code will never be executed because the default max batch size is 1000.
                        Console.WriteLine("The input request collection contains %0 requests, which exceeds the maximum allowed (%1)",
                            transactionRequest.Requests.Count, maxBatchSize);
                        Log(string.Format("The input request collection contains %0 requests, which exceeds the maximum allowed (%1)",
                            transactionRequest.Requests.Count, maxBatchSize), log);
                    }
                }
                else
                {
                    Console.WriteLine("Transaction rolled back because: {0}", fault.Message);
                    Log(string.Format("Transaction rolled back because: {0}",
                           fault.Message), log);
                }
                // Re-throw so Main() can process the fault.
                throw;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine("The application terminated with an error.");
                Console.WriteLine(ex.Message);
                Log("The application terminated with an error. " + ex.Message, log);
            }
        }

        /// <summary>
        /// This function log the data to file
        /// </summary>
        /// <param name="logMessage"></param>
        /// <param name="w"></param>
        public static void Log(string logMessage, TextWriter w)
        {
            w.Write("\r\n");
            w.WriteLine($" {logMessage}");
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


        #region CRM CONNECTION
        public static IOrganizationService CreateService(string serverUrl, string userId, string password, string domain)
        {
            Console.WriteLine("\n\nConnecting to CRM..........\n\n");

            //objDataValidation.CreateLog("Before CRm  Creation");
            IOrganizationService _service;

            ClientCredentials Credentials = new ClientCredentials();
            ClientCredentials devivceCredentials = new ClientCredentials();

           // Credentials.UserName.UserName = domain + "\\" + userId;
            Credentials.UserName.UserName = userId;
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

        #endregion
    }
    #endregion 
}
#endregion
