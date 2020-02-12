using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.ServiceModel.Description;
using System.Net;
using Microsoft.Xrm.Sdk.Client;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

namespace NeglectedOppsCloseTool
{
    class Program
    {
        static void Main(string[] args)
        {
            string Progress = "Start Neglected Opps Close Tool";
            
            try
            {
                #region Setup DB Connections
                //CRM DB - Change connection to UAT or Live
                CrmConnection con = new CrmConnection("ARUP_CRM");
                //CrmConnection con = new CrmConnection("CRM_Live");
                Progress = "Connection to CRM made";
                
                //Housekeeping DB - Change connection to UAT or Live
                SqlConnection HKconn = new SqlConnection(ConfigurationManager.ConnectionStrings["CRM_Environment"].ConnectionString);
                //SqlConnection HKconn = new SqlConnection(ConfigurationManager.ConnectionStrings["HouseKeeping_Live"].ConnectionString);                 
                //SqlConnection HKconn = new SqlConnection( ConfigurationManager.ConnectionStrings["HouseKeeping_Chee"].ConnectionString);
                Progress = "Connection to Housekeeping DB made";

                #endregion

                #region Find Neglected Opportunities to Close
              
                string fetchQuery = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='opportunity'>
                    <attribute name='name' />
                    <attribute name='opportunityid' />
                    <attribute name='ccrm_reference' />
                    <attribute name='modifiedon' />
                    <attribute name='ccrm_closeneglectedopportunitymarkdate' />
                    <attribute name='ccrm_jna' />
                    <attribute name='estimatedvalue_base' />
                    <order attribute='name' descending='false' />
                    <filter type='and'>
                      <condition attribute='statecode' operator='eq' value='0' />
                      <condition attribute='ccrm_closeneglectedopportunity' operator='eq' value='1' />
                    </filter>
                  </entity>
                </fetch>";
                //IOrganizationService service = new OrganizationService(con);  
                string serverUrl = ConfigurationManager.AppSettings["serverUrl"];
                string userId = ConfigurationManager.AppSettings["userId"];
                string password = ConfigurationManager.AppSettings["password"];
                string domain = ConfigurationManager.AppSettings["domain"];
                //IOrganizationService service = CreateService("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc", "crm.hub@arup.com", "CIm2$98pRt", "arup");
                IOrganizationService service = CreateService(serverUrl, userId, password, domain);
                //Implemented loop to handle the 5000 record limit
                while (true)
                {
                    EntityCollection fetchResults = service.RetrieveMultiple(new FetchExpression(fetchQuery));
                    Entity opportunityClose = new Entity("opportunityclose");
                    Progress = "Retrieved neglected opportunities";

                    //Open conenction to Housekeeping DB if opportunities to close
                    if (fetchResults.Entities.Count > 0)
                    {
                        
                        //HKconn.Open();
                        var updateOptyError = "";
                        foreach (var oppo in fetchResults.Entities)
                        {
                            #region Close Neglected Opportunity
                            try
                            {
                                String closeType = "";
                                
                                Console.WriteLine("Opportunity" + " " + oppo.GetAttributeValue("opportunityid") + " " + oppo.GetAttributeValue("name"));

                                Progress = "Processing opportunity: " + oppo.GetAttributeValue("opportunityid") + " " + oppo.GetAttributeValue("name");

                                //If not modified since marked to be closed then close opportunity - minus one min off modified date in case took a while to save
                                if ((DateTime)oppo["ccrm_closeneglectedopportunitymarkdate"] >= Convert.ToDateTime(oppo["modifiedon"]).AddMinutes(-1))
                                {
                                    //Close Opportunity as Won if CJN assigned                    
                                    if (oppo.Contains("ccrm_jna"))
                                    {
                                        closeType = "Won";
                                        WinOpportunityRequest winReq = new WinOpportunityRequest();

                                        opportunityClose["opportunityid"] = new EntityReference("opportunity", oppo.GetAttributeValue<Guid>("opportunityid"));
                                        opportunityClose["actualend"] = DateTime.Now;
                                        opportunityClose["description"] = "Opportunity automatically closed as Won by Neglected Opportunities process ";
                                        opportunityClose["actualrevenue"] = oppo["estimatedvalue_base"];

                                        winReq.OpportunityClose = opportunityClose;
                                        winReq.Status = new OptionSetValue(3);
                                        WinOpportunityResponse resp = (WinOpportunityResponse)service.Execute(winReq);
                                    }
                                    //Close Opportunity as Lost if no CJN assigned                    
                                    else
                                    {
                                        closeType = "Lost";
                                        LoseOpportunityRequest loseReq = new LoseOpportunityRequest();

                                        opportunityClose["opportunityid"] = new EntityReference("opportunity", oppo.GetAttributeValue<Guid>("opportunityid"));
                                        opportunityClose["subject"] = "Automatically closed - Opportunity inactive!";
                                        opportunityClose["description"] = "Opportunity automatically closed based on: Lead not modified in the last 4 months or Arup Project Start date missing or earlier than today ";
                                        opportunityClose["actualend"] = DateTime.Now;
                                        loseReq.OpportunityClose = opportunityClose;

                                        loseReq.Status = new OptionSetValue(100000005);
                                        LoseOpportunityResponse resp = (LoseOpportunityResponse)service.Execute(loseReq);
                                    }

                                    ////Log in Housekeeping DB                        
                                    //using (SqlCommand cmd = new SqlCommand("INSERT INTO NeglectedOpps_Closed (CloseDate, OpportunityID, ProjectID, CloseType) VALUES (@CloseDate, @OpportunityID, @ProjectID, @CloseType)", HKconn))
                                    //{
                                    //    cmd.Connection = HKconn;
                                    //    cmd.CommandType = CommandType.Text;
                                    //    cmd.Parameters.AddWithValue("@CloseDate", DateTime.Now);
                                    //    cmd.Parameters.AddWithValue("@OpportunityID", oppo.GetAttributeValue<Guid>("opportunityid"));
                                    //    cmd.Parameters.AddWithValue("@ProjectID", oppo["ccrm_reference"].ToString());
                                    //    cmd.Parameters.AddWithValue("@CloseType", closeType);
                                    //    cmd.ExecuteNonQuery();
                                    //}

                                }

                                //Reset Close Opportunity Flags in Opportunity Record             
                                Entity updatedOppo = new Entity("opportunity");
                                updatedOppo.Id = oppo.GetAttributeValue<Guid>("opportunityid");
                                updatedOppo["ccrm_closeneglectedopportunity"] = null;
                                updatedOppo["ccrm_closeneglectedopportunitymarkdate"] = null;

                                // Update the opportunity.                
                                service.Update(updatedOppo);
                            }
                            catch (Exception e)
                            {
                                updateOptyError = updateOptyError + System.Environment.NewLine + "Error with Close Neglected Opportunities: Progress='" + Progress + "'.  Error: " + e.Message;

                            }

                            
                            #endregion
                        }
                        if (updateOptyError != null && updateOptyError != "")
                        {
                            Log(updateOptyError);
                        }
                        //Close conenction to Housekeeping DB if open
                        if (HKconn.State == System.Data.ConnectionState.Open) HKconn.Close();
                                                  
                    }
                    else
                    {
                        // If no more records in the result nodes, exit the loop.
                        break;
                    }                                        
                }
                
                #endregion                 

            }
            catch (Exception ex)
            {
                Log("Error with Close Neglected Opportunities: Progress='" + Progress + "'.  Error: " + ex.Message);

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

        public static void Log(string logMessage)
        {            
            TextWriter w = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + @"CloseNeglectedOppsErrorLog.txt", true);

            w.Write("\r\nLog Entry : ");
            w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());           
            w.WriteLine(":{0}", logMessage);
            w.WriteLine("-------------------------------");
            w.Close();
        }
        
    }
}
