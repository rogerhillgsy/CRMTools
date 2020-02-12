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

namespace Framework
{
    class Program
    {
        static List<string> linesInFailedFile = null;
        static void Main(string[] args)
        {
            try
            {
                IOrganizationService service = CreateService("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc", "crm.hub@arup.com", "CIm2$98pRt", "arup");
                UpdateOpportunity(service);
                UpdateOpportunity(service);
            }
            catch (Exception ex)
            {
                //Console.WriteLine("Error occured : ", ex.InnerException.Message);
                Console.WriteLine("Error occured : {0} ", ex.Message);
            }
            finally
            {
                System.IO.File.WriteAllLines("FailedRecordsFile" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".csv", linesInFailedFile);
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
            query.PageInfo.Count = 5;
            query.PageInfo.PageNumber = 1;
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
            }
            while (entityCollection.MoreRecords);
            Console.WriteLine("Total Framework record count:" + final.TotalRecordCount);
            Console.ReadKey();
        }

        //ccrm_othernetworksval", "ccrm_servicesvalue", "ccrm_theworksvalue", "ccrm_disciplinesvalue", "ccrm_projectsectorvalue"
        public static void UpdateFrameworkMultiSelect(IOrganizationService service, Guid arup_frameworkid, string arup_arupregionvalue, string arup_disciplinesvalue, string arup_projecttypevalue, string arup_projectsectorvalue, string arup_servicesvalue)
        {
            try
            {
                Entity opportunity = new Entity("arup_framework");
                if (arup_arupregionvalue != string.Empty && arup_arupregionvalue != null)
                {
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_arupregionvalue.Split(',');
                    foreach (var item in arr)
                    {
                        collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                    }

                    opportunity["arup_region"] = collectionOptionSetValues;
                }
                if (arup_disciplinesvalue != string.Empty && arup_disciplinesvalue != null)
                {
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_disciplinesvalue.Split(',');
                    foreach (var item in arr)
                    {
                        collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                    }

                    opportunity["arup_disciplines_ms"] = collectionOptionSetValues;
                }
                if (arup_projecttypevalue != string.Empty && arup_projecttypevalue != null)
                {
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_projecttypevalue.Split(',');
                    foreach (var item in arr)
                    {
                        collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                    }

                    opportunity["arup_projecttype_ms"] = collectionOptionSetValues;
                }
                if (arup_projectsectorvalue != string.Empty && arup_projectsectorvalue != null)
                {
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_projectsectorvalue.Split(',');
                    foreach (var item in arr)
                    {
                        collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                    }

                    opportunity["arup_projectsector_ms"] = collectionOptionSetValues;
                }
                if (arup_servicesvalue != string.Empty && arup_servicesvalue != null)
                {
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = arup_servicesvalue.Split(',');
                    foreach (var item in arr)
                    {
                        collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
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
    }
}
