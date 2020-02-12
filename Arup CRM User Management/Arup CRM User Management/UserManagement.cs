using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel.Description;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
//using CrmEarlyBound;

//********************** Batch Job for Product Backlog Item 33479 to manage CRM user licenses ********************

namespace Arup_CRM_User_Management
{
    class UserManagement
    {
        private static IOrganizationService _orgService;
        private static IOrganizationService _orgTrainingService;
        static void Main(string[] args)
        {
            try
            {
                CrmConnection connection = CrmConnection.Parse(
                ConfigurationManager.ConnectionStrings["CRMConnectionString"].ConnectionString);
                CrmConnection connectionTraining = CrmConnection.Parse(ConfigurationManager.ConnectionStrings["CRMConnectionTrainingString"].ConnectionString);
                List<string> errorList = new List<string>(); //List to store error for each user record
                int inactivitydays = int.Parse(ConfigurationManager.AppSettings["InactivityCheckDays"]);
                DateTime createdOnChkDt = DateTime.Now.AddDays(-inactivitydays);
                int phase1days = int.Parse(ConfigurationManager.AppSettings["Phase1Days"]);
                int reminderdays = int.Parse(ConfigurationManager.AppSettings["ReminderDays"]);
                string execlogpath = ConfigurationManager.AppSettings["ExecutionLogPath"];
                string errorlogpath = ConfigurationManager.AppSettings["ErrorLogPath"];
                string executiondttime = DateTime.Now.ToString("yyyyMMddTHHmmss");
                int pageSize = 5000;
                int pageNumber = 1;
                string pagingCookie = string.Empty;
                int totalRecordCount = 0;
                //_orgService = CreateService("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc", "crm.hub@arup.com", "CIm2$98pRt", "arup");
                string serverUrl = ConfigurationManager.AppSettings["serverUrl"];
                string userId = ConfigurationManager.AppSettings["usersId"];
                string password = ConfigurationManager.AppSettings["password"];
                string domain = ConfigurationManager.AppSettings["domain"];
                IOrganizationService _orgService = CreateService(serverUrl, userId, password, domain);
                //using (_orgService = new OrganizationService(connection))
                //{
                // Log the execution start to Execution Log
                //File.WriteAllText(execlogpath + @"\Arup User Management Execution Log_" + executiondttime + ".txt", "***************Arup User Management Batch Job***************" + Environment.NewLine);
                WriteLog("***************Arup User Management Batch Job***************" + Environment.NewLine, false);
                //File.AppendAllText(execlogpath + @"\Arup User Management Execution Log_" + executiondttime + ".txt", "Arup User Management Batch Job execution started at:" + DateTime.Now + Environment.NewLine);
                WriteLog("Arup User Management Batch Job execution started at:" + DateTime.Now + Environment.NewLine + Environment.NewLine, false);

                Guid _systemUserId; //For testing only

                _systemUserId = Guid.Parse(ConfigurationManager.AppSettings["UserId"]);//For testing only

                #region Query Users
                //Query Active users in the system and retreive their System User ID               
                var userquery = new QueryExpression(SystemUser.EntityLogicalName)
                {
                    ColumnSet = new ColumnSet("systemuserid", "ccrm_phase1deactivationdate", "ccrm_phase2deactivationdate", "ccrm_deactivationreminderdate", "domainname"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };
                //Add condition in the query to check if user is enabled in the system
                userquery.Criteria.AddCondition(
                  "isdisabled", ConditionOperator.Equal, false);
                userquery.Criteria.AddCondition(
                  "ccrm_donotdeactivate", ConditionOperator.NotEqual, true);
                userquery.Criteria.AddCondition(
                  "createdon", ConditionOperator.LessEqual, createdOnChkDt);
                //userquery.Criteria.AddCondition(
                //  "systemuserid", ConditionOperator.Equal, _systemUserId);

                // Assign the pageinfo properties to the query expression.
                userquery.PageInfo = new PagingInfo();
                userquery.PageInfo.PageNumber = pageNumber;
                userquery.PageInfo.Count = pageSize;
                userquery.PageInfo.PagingCookie = pagingCookie;
                try
                {
                    while (true)
                    {
                        var userresults = _orgService.RetrieveMultiple(userquery);

                        totalRecordCount += userresults.Entities.Count;


                        #region User Access Check
                        // For each user record check if they have logged within the last logon cut off date
                        foreach (SystemUser user in userresults.Entities)
                        {
                            try
                            {
                                //Check if the user has been activated during the last 90 days

                                EntityCollection activationresults = new EntityCollection();
                                activationresults = QueryLastActivation(user.SystemUserId, DateTime.Now.AddDays(-inactivitydays));//Change to hours for testing

                                if (activationresults.Entities.Count == 0)
                                {
                                    //Check if Phase 1 Date is populated on user record. 
                                    //RS[06/02/2017] - This field will now used to store deactivation date
                                    //as Phase wise deactivation has been ruled out by business
                                    if (user.ccrm_Phase1DeactivationDate != null)
                                    {

                                        //Check if the deactivation date is today    
                                        if (user.ccrm_Phase1DeactivationDate <= DateTime.Now)
                                        {
                                            //Query Audit table to check if user has accessed the system in last 120 days

                                            EntityCollection auditresults = new EntityCollection();
                                            auditresults = QueryAudit(user.SystemUserId, DateTime.Now.AddDays(-(phase1days + inactivitydays)));//change to hours for testing
                                                                                                                                               // training user based on userid

                                            if (auditresults.Entities.Count == 0)
                                            {
                                                //user.ccrm_Phase1DeactivationDate = null;
                                                user.ccrm_SystemDisabled = true;
                                                _orgService.Update(user);

                                                SetStateRequest disableusrrequest = new SetStateRequest()
                                                {
                                                    EntityMoniker = user.ToEntityReference(),
                                                    // Sets the user to disabled.
                                                    State = new OptionSetValue(1),
                                                    // Required by request but always valued at -1 in this context.
                                                    Status = new OptionSetValue(-1)

                                                };
                                                _orgService.Execute(disableusrrequest);

                                                // disabling user in training 
                                                //using (_orgTrainingService = new OrganizationService(connectionTraining))
                                                //{
                                                DisableUserInTraining(_orgService, user);
                                                //}

                                            }
                                            else
                                            {
                                                user.ccrm_Phase1DeactivationDate = null;
                                                user.ccrm_Phase1NotificationDate = null;
                                                user.ccrm_DeactivationReminderDate = null;
                                                _orgService.Update(user);
                                            }

                                        }//End of If Deactivation date is today or in past
                                        else
                                        {
                                            //Check if the Deactivation Date is 7 days from now
                                            if ((user.ccrm_Phase1DeactivationDate.Value.Subtract(TimeSpan.FromDays(reminderdays)) <= DateTime.Now) && user.ccrm_DeactivationReminderDate == null)
                                            {
                                                //Query Audit table to check if user has accessed the system in last 90 days i.e reminder days before deactivation date.

                                                EntityCollection auditresults = new EntityCollection();
                                                //auditresults = QueryAudit(user.SystemUserId, user.ccrm_Phase1DeactivationDate.Value.Subtract(TimeSpan.FromDays(inactivitydays)));//change to hours for testing
                                                auditresults = QueryAudit(user.SystemUserId, DateTime.Now.AddDays(-inactivitydays));//change to hours for testing
                                                if (auditresults.Entities.Count == 0)
                                                {
                                                    //If the user has not accessed the system then set the reminder date to trigger a reminder email from the workflow
                                                    user.ccrm_DeactivationReminderDate = DateTime.Today;
                                                    _orgService.Update(user);

                                                }
                                                else
                                                {
                                                    //If the user has accessed the system then nullify the flags
                                                    user.ccrm_Phase1DeactivationDate = null;
                                                    user.ccrm_DeactivationReminderDate = null;
                                                    user.ccrm_Phase1NotificationDate = null;
                                                    _orgService.Update(user);
                                                }

                                            }
                                        }

                                    }
                                    //End of If Deactivation Date is present on the user record.
                                    else
                                    {
                                        EntityCollection auditresults = new EntityCollection();
                                        auditresults = QueryAudit(user.SystemUserId, DateTime.Now.AddDays(-inactivitydays));//Change to hours for testing
                                        if (auditresults.Entities.Count == 0)
                                        {
                                            user.ccrm_Phase1DeactivationDate = DateTime.Today.AddDays(phase1days);//Comment for testing
                                                                                                                  // user.ccrm_Phase1DeactivationDate = DateTime.Now.AddHours(phase1days);//Uncomment for testing
                                            user.ccrm_Phase1NotificationDate = DateTime.Today; //Set the notification date to today. Used in the email sent out to users.

                                            _orgService.Update(user);

                                        }
                                        else
                                        {
                                            // Do nothing as user is an active user
                                        }
                                    }
                                }// End user activation check if

                            }//End try within for loop
                            catch (Exception auditcheckex)
                            {
                                string auditcheckerror = "Error processing user record:" + user.FullName + auditcheckex.Message;
                                errorList.Add(auditcheckerror);
                            }

                        }//End of Users For loop
                        #endregion User Access Check

                        if (userresults.MoreRecords)
                        {
                            // Increment the page number to retrieve the next page.
                            userquery.PageInfo.PageNumber++;

                            // Set the paging cookie to the paging cookie returned from current results.
                            userquery.PageInfo.PagingCookie = userresults.PagingCookie;
                        }
                        else
                        {
                            // If no more records are in the result nodes, exit the loop.
                            break;
                        }


                    }//End of while loop
                     // Append User Query results to Execution Log file
                     //File.AppendAllText(execlogpath + @"\Arup User Management Execution Log_" + executiondttime + ".txt", "Retreived user records:" + totalRecordCount + Environment.NewLine);
                    WriteLog("Retreived user records:" + totalRecordCount + Environment.NewLine, false);
                    #endregion Query Users
                }
                catch (Exception userqueryex)
                {
                    string userqueryerror = "Error querying the user records:" + userqueryex.Message;
                    errorList.Add(userqueryerror);
                }


                #region Error Logging
                if (errorList.Any())
                {
                    // Write list of error user records to Error Log File
                    string[] errorarray = errorList.ToArray();
                    //Create Error Log File
                    //File.WriteAllText(errorlogpath + @"\Arup User Management Error Log_" + executiondttime + ".txt", "***************Arup User Management Batch Job Error Details***************" + Environment.NewLine);
                    WriteLog("***************Arup User Management Batch Job Error Details***************" + Environment.NewLine, true);
                    //Append List of Error records to Error Log file
                    //File.AppendAllLines(errorlogpath + @"\Arup User Management Error Log_" + executiondttime + ".txt", errorarray);
                    foreach (string line in errorarray)
                        WriteLog(line, true);
                    // Write to Execution Log file that there is an error
                    //File.AppendAllText(execlogpath + @"\Arup User Management Execution Log_" + executiondttime + ".txt", "Job completed at:" + DateTime.Now + " with errors, please refer error log file.");
                    WriteLog("Job completed at:" + DateTime.Now + " with errors, please refer error log file.", true);
                }
                else
                {
                    // Write to Execution Log file that job completed successfully
                    // File.AppendAllText(execlogpath + @"\Arup User Management Execution Log_" + executiondttime + ".txt", "Job completed successfully at:" + DateTime.Now);
                    WriteLog("Job completed successfully at:" + DateTime.Now, false);
                }
                #endregion Error Logging
                //}
            }//End try
            catch (FaultException<OrganizationServiceFault> ex)
            {
                string message = ex.Message;
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
        private static void DisableUserInTraining(IOrganizationService _orgService, SystemUser user)
        {
            // find the user in training
            var trainingUser = Guid.Empty;
            QueryExpression trainingExp = new QueryExpression(SystemUser.EntityLogicalName);
            trainingExp.ColumnSet = new ColumnSet();
            trainingExp.Criteria.AddCondition("domainname", ConditionOperator.Equal, user.DomainName);
            var results = _orgTrainingService.RetrieveMultiple(trainingExp);
            if (results.Entities.Count > 0)
                foreach (var entity in results.Entities)
                {
                    trainingUser = entity.Id;
                }
            if (!Guid.Equals(trainingUser, Guid.Empty))
            {
                var Trainingentity = new SystemUser();
                Trainingentity.Id = trainingUser;
                Trainingentity.ccrm_SystemDisabled = true;
                _orgTrainingService.Update(Trainingentity);

                SetStateRequest disableusrTrainingrequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference(SystemUser.EntityLogicalName, trainingUser),
                    // Sets the user to disabled.
                    State = new OptionSetValue(1),
                    // Required by request but always valued at -1 in this context.
                    Status = new OptionSetValue(-1)

                };
                _orgTrainingService.Execute(disableusrTrainingrequest);

            }
        }

        public static EntityCollection QueryAudit(Guid? userid, DateTime lstlogondt)
        {
            //Method to query Audit table based on the userid and timeframe provided by the main function
            //Returns the entity collection results from Audit table
            try
            {

                //Query Audit table for user access records
                var query = new QueryExpression(Audit.EntityLogicalName)
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };

                // Only retrieve audit records that track user access.
                query.Criteria.AddCondition("action", ConditionOperator.In,
                    //(int)AuditAction.UserAccessviaWebServices,
                    (int)AuditAction.UserAccessviaWeb);
                //Add filter condition for retreived user and last logon cut off date
                var userFilter = query.Criteria.AddFilter(LogicalOperator.And);
                userFilter.AddCondition(
                    "objectid", ConditionOperator.Equal, userid);
                userFilter.AddCondition(
                  "createdon", ConditionOperator.GreaterThan, lstlogondt);

                var results = _orgService.RetrieveMultiple(query);
                return (results);
            }
            catch (Exception e)
            {
                //Throw Error to main function
                throw (e);
            }
        }

        public static EntityCollection QueryLastActivation(Guid? userid, DateTime lstactivationdt)
        {
            //Method to query Audit table based on the userid and timeframe of last activation provided by the main function
            //Returns the entity collection results from Audit table 
            try
            {

                //Query Audit table for user access records
                var query = new QueryExpression(Audit.EntityLogicalName)
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression(LogicalOperator.And)
                };

                //Only retrieve audit records that track user access.
                query.Criteria.AddCondition("action", ConditionOperator.In,
                    //(int)AuditAction.UserAccessviaWebServices,
                    (int)AuditAction.Activate);
                //Add filter condition for retreived user and last logon cut off date
                var userFilter = query.Criteria.AddFilter(LogicalOperator.And);
                userFilter.AddCondition(
                    "objectid", ConditionOperator.Equal, userid);
                userFilter.AddCondition(
                  "createdon", ConditionOperator.GreaterThan, lstactivationdt);

                var results = _orgService.RetrieveMultiple(query);
                return (results);
            }
            catch (Exception e)
            {
                //Throw Error to main function
                throw (e);
            }

        }

        public static void WriteLog(string strLog, bool errorlog)
        {
            StreamWriter log;
            FileStream fileStream = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;
            string logFilePath;

            //string logFilePath = ConfigurationManager.AppSettings["LogPath"].ToString();
            if (errorlog)
                logFilePath = ConfigurationManager.AppSettings["ExecutionLogPath"] + @"\Error Log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
            else
                logFilePath = ConfigurationManager.AppSettings["ExecutionLogPath"] + @"\Execution Log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            logFileInfo = new FileInfo(logFilePath);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists) logDirInfo.Create();
            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);
            }
            log = new StreamWriter(fileStream);
            log.WriteLine(strLog);
            log.Close();
        }

    }
}
