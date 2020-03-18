﻿using Microsoft.Crm.Sdk.Messages;
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

namespace Framework
{
    class Program
    {
        static List<string> linesInFailedFile = null;
        static string fileName = "FailedRecordsFile" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".csv";
        static int StartPageNumber = 0;
        static int EndPageNumber = 0;
        static int RecordCountPerPage = 0;
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Framework records Processing Statred. Start time:" + DateTime.Now);
                linesInFailedFile = new List<string>();
                linesInFailedFile.Add("Entity,RecordId,Error Description, OptionSetValues");
                linesInFailedFile.Add(string.Format("{0},{1},{2},{3}", "Start Time : " + DateTime.Now, "", "", ""));
                string serverUrl = ConfigurationManager.AppSettings["serverUrl"].ToString();
                string userName = ConfigurationManager.AppSettings["UserName"].ToString();
                string password = ConfigurationManager.AppSettings["Password"].ToString();
                string domain = ConfigurationManager.AppSettings["Domain"].ToString();
                StartPageNumber = Convert.ToInt32(ConfigurationManager.AppSettings["StartPageNumber"]);
                EndPageNumber = Convert.ToInt32(ConfigurationManager.AppSettings["EndPageNumber"]);
                RecordCountPerPage = Convert.ToInt32(ConfigurationManager.AppSettings["RecordCountPerPage"]);
                IOrganizationService service = CreateService(serverUrl, userName, password, domain);
                //IOrganizationService service = CreateService("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc", "crm.hub@arup.com", "CIm2$98pRt", "arup");
                UpdateOpportunity(service);
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
        public static void UpdateOpportunity(IOrganizationService service)
        {
            QueryExpression test = new QueryExpression();

            QueryExpression query = new QueryExpression("arup_framework");
            //query.ColumnSet = new ColumnSet("ccrm_businessinterestpicklistname", "ccrm_businessinterestpicklistvalue", "arup_businessinterest");
            query.ColumnSet.AddColumns("arup_arupregionvalue", "arup_disciplinesvalue", "arup_projecttypevalue", "arup_projectsectorvalue", "arup_servicesvalue");
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.Or;
            query.Criteria.AddCondition("arup_arupregionvalue", ConditionOperator.NotNull);
            query.Criteria.AddCondition("arup_disciplinesvalue", ConditionOperator.NotNull);
            query.Criteria.AddCondition("arup_projecttypevalue", ConditionOperator.NotNull);
            query.Criteria.AddCondition("arup_projectsectorvalue", ConditionOperator.NotNull);
            query.Criteria.AddCondition("arup_servicesvalue", ConditionOperator.NotNull);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = RecordCountPerPage;
            query.PageInfo.PageNumber = StartPageNumber;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            EntityCollection final = new EntityCollection();
            foreach (Entity i in entityCollection.Entities)
            {
                final.Entities.Add(i);
                UpdateFrameworkMultiSelect(service,
                    i.GetAttributeValue<Guid>("arup_frameworkid"),
                    i.GetAttributeValue<string>("arup_arupregionvalue"),
                    i.GetAttributeValue<string>("arup_disciplinesvalue"),
                    i.GetAttributeValue<string>("arup_projecttypevalue"),
                    i.GetAttributeValue<string>("arup_projectsectorvalue"),
                    i.GetAttributeValue<string>("arup_servicesvalue"));
            }
            do
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                entityCollection = service.RetrieveMultiple(query);
                foreach (Entity i in entityCollection.Entities)
                {
                    final.Entities.Add(i);
                    UpdateFrameworkMultiSelect(service,
                    i.GetAttributeValue<Guid>("arup_frameworkid"),
                    i.GetAttributeValue<string>("arup_arupregionvalue"),
                    i.GetAttributeValue<string>("arup_disciplinesvalue"),
                    i.GetAttributeValue<string>("arup_projecttypevalue"),
                    i.GetAttributeValue<string>("arup_projectsectorvalue"),
                    i.GetAttributeValue<string>("arup_servicesvalue"));
                }
                Console.WriteLine(query.PageInfo.PageNumber * RecordCountPerPage + " Framework Records processed at : " + DateTime.Now);
                System.IO.File.WriteAllLines(fileName, linesInFailedFile);
                if (query.PageInfo.PageNumber == EndPageNumber)
                    break;
            }
            while (entityCollection.MoreRecords);
            Console.WriteLine("Total Framework record count:" + RecordCountPerPage * EndPageNumber);
            Console.WriteLine("Framework records Processing Completed. End time:" + DateTime.Now);
            Console.ReadKey();
        }

        //ccrm_othernetworksval", "ccrm_servicesvalue", "ccrm_theworksvalue", "ccrm_disciplinesvalue", "ccrm_projectsectorvalue"
        public static void UpdateFrameworkMultiSelect(IOrganizationService service, Guid arup_frameworkid, string arup_arupregionvalue, string arup_disciplinesvalue, string arup_projecttypevalue, string arup_projectsectorvalue, string arup_servicesvalue)
        {
            try
            {
                //arup_arupregionvalue = ",, , ,100000009,";
                Entity opportunity = new Entity("arup_framework");
                if (arup_arupregionvalue != string.Empty && arup_arupregionvalue != null)
                {
                    Dictionary<Nullable<int>, string> opset = RetriveOptionSetLabels(service, "arup_framework", "arup_arupregionpicklist");
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_arupregionvalue.Split(',');
                    foreach (var item in arr)
                    {
                        if (item != null && item.Trim() != string.Empty && item.Trim() != "")
                        {
                            if (opset.ContainsKey(Convert.ToInt32(item)))
                            {
                                collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                            }
                        }
                    }

                    opportunity["arup_region"] = collectionOptionSetValues;
                }
                if (arup_disciplinesvalue != string.Empty && arup_disciplinesvalue != null)
                {
                    Dictionary<Nullable<int>, string> opset = RetriveOptionSetLabels(service, "arup_framework", "arup_disciplinespicklist");
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_disciplinesvalue.Split(',');
                    foreach (var item in arr)
                    {
                        if (item != null && item.Trim() != string.Empty && item.Trim() != "")
                        {
                            if (opset.ContainsKey(Convert.ToInt32(item)))
                            {
                                collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                            }
                        }
                    }

                    opportunity["arup_disciplines_ms"] = collectionOptionSetValues;
                }
                if (arup_projecttypevalue != string.Empty && arup_projecttypevalue != null)
                {
                    Dictionary<Nullable<int>, string> opset = RetriveOptionSetLabels(service, "arup_framework", "arup_projecttypepicklist");
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_projecttypevalue.Split(',');
                    foreach (var item in arr)
                    {
                        if (item != null && item.Trim() != string.Empty && item.Trim() != "")
                        {
                            if (opset.ContainsKey(Convert.ToInt32(item)))
                            {
                                collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                            }
                        }
                    }

                    opportunity["arup_projecttype_ms"] = collectionOptionSetValues;
                }
                if (arup_projectsectorvalue != string.Empty && arup_projectsectorvalue != null)
                {
                    Dictionary<Nullable<int>, string> opset = RetriveOptionSetLabels(service, "arup_framework", "arup_projectsectorpicklist");
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_projectsectorvalue.Split(',');
                    foreach (var item in arr)
                    {
                        if (item != null && item.Trim() != string.Empty && item.Trim() != "")
                        {
                            if (opset.ContainsKey(Convert.ToInt32(item)))
                            {
                                if (opset.ContainsKey(Convert.ToInt32(item)))
                                {
                                    collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                                }
                            }
                        }
                    }

                    opportunity["arup_projectsector_ms"] = collectionOptionSetValues;
                }
                if (arup_servicesvalue != string.Empty && arup_servicesvalue != null)
                {
                    Dictionary<Nullable<int>, string> opset = RetriveOptionSetLabels(service, "arup_framework", "arup_servicespicklist");
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_servicesvalue.Split(',');
                    foreach (var item in arr)
                    {
                        if (item != null && item.Trim() != string.Empty && item.Trim() != "")
                        {
                            collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item))); if (opset.ContainsKey(Convert.ToInt32(item)))
                            {
                                if (opset.ContainsKey(Convert.ToInt32(item)))
                                {
                                    collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                                }
                            }
                        }
                    }

                    opportunity["arup_services_ms"] = collectionOptionSetValues;
                }
                opportunity.Id = arup_frameworkid;
                service.Update(opportunity);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error : " + e.Message);
                //linesInFailedFile.Add("RecordId,ToOptionset,Values,Message");
                string optionSetValues = "arup_region : " + arup_arupregionvalue + " | arup_disciplines_ms : " + arup_disciplinesvalue + " | arup_projecttype_ms : " + arup_projecttypevalue + " | arup_projectsector_ms : " + arup_projectsectorvalue + " | arup_services_ms : " + arup_servicesvalue;

                linesInFailedFile.Add(string.Format("{0},{1},{2},{3}", "arup_framework", arup_frameworkid, e.Message, optionSetValues));
            }
        }
        #endregion

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
            FetchXml = FetchXml + "<condition attribute='objecttypecodename' operator='eq' value='" + EntityLogicalName + "' />";
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
    }
}
