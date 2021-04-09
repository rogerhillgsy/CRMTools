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
using Microsoft.Xrm.Tooling.Connector;

namespace NeglectedOppsCloseTool
{
    class Program
    {
        static List<string> logEntry = null;
        static string fileName = "";
        static void Main(string[] args)
        {
            string Progress = "Start Neglected Opps Close Tool";
            string logFilePath = ConfigurationManager.AppSettings.Get("LogFilePath");

            fileName = logFilePath + "CloseNeglectedOppsErrorLog_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".txt";
            try
            {
                #region Setup DB Connections
                //CRM DB - Change connection to UAT or Live
                //CrmConnection con = new CrmConnection("ARUP_CRM");
                //CrmConnection con = new CrmConnection("CRM_Live");
                Progress = "Connection to CRM made";
                
                //Housekeeping DB - Change connection to UAT or Live
                //SqlConnection HKconn = new SqlConnection(ConfigurationManager.ConnectionStrings["CRM_Environment"].ConnectionString);
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

                logEntry = new List<string>();
                logEntry.Add(string.Format("{0}", "Start Time : " + DateTime.Now));

                Console.WriteLine("\n\nConnecting to CRM..........\n\n");
                logEntry.Add(string.Format("{0}", "\n\nConnecting to CRM..........\n"));
                //Log("\n\nConnecting to CRM..........\n\n", logTxtWriter);
                IOrganizationService service;
                string password1 = "CIm2$98pRt";
                string ConnectionString = ConfigurationManager.ConnectionStrings["CrmCloudConnection"].ConnectionString;
                ConnectionString = ConnectionString.Replace("%Password%", password1);
                //Console.WriteLine("ConnectionString is ::" + ConnectionString);
                var CrmService = new CrmServiceClient(ConnectionString);
                service = CrmService.OrganizationServiceProxy;
                Console.WriteLine("Connection Established!!!\n\n");
                //Log("\n\nConnection Established!!!\n\n", logTxtWriter);
                logEntry.Add(string.Format("{0}", "\n\nConnection Established..........\n"));
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
                                //Console.WriteLine("Opportunity" + " " + oppo.GetAttributeValue("opportunityid") + " " + oppo.GetAttributeValue("name"));
                                logEntry.Add(string.Format("{0}", "Opportunity" + " " + oppo.GetAttributeValue("opportunityid") + " " + oppo.GetAttributeValue("name")));
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
                            //Log(updateOptyError, logTxtWriter);
                            logEntry.Add(string.Format("{0}", updateOptyError));
                        }
                        //Close conenction to Housekeeping DB if open
                        //if (HKconn.State == System.Data.ConnectionState.Open) HKconn.Close();                                                  
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
                logEntry.Add(string.Format("{0}", "Error with Close Neglected Opportunities: Progress='" + Progress + "'.  Error: " + ex.Message));                
                System.IO.File.WriteAllLines(fileName, logEntry);
            }
            finally
            {
                logEntry.Add(string.Format("{0}", "End Time='" + DateTime.Now));
                System.IO.File.WriteAllLines(fileName, logEntry);
            }
        }      
    }
}
