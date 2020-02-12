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

namespace Organisation
{
    class Program
    {
        static void Main(string[] args)
        {
            IOrganizationService service = CreateService("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc", "crm.hub@arup.com", "CIm2$98pRt", "arup");
            UpdateOrganisation(service);
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

        #region Update Organisation
        public static void UpdateOrganisation(IOrganizationService service)
        {
            QueryExpression test = new QueryExpression();

            QueryExpression query = new QueryExpression("opportunity");
            //query.ColumnSet = new ColumnSet("ccrm_businessinterestpicklistname", "ccrm_businessinterestpicklistvalue", "arup_businessinterest");
            query.ColumnSet.AddColumns("ccrm_othernetworksval", "ccrm_servicesvalue", "ccrm_theworksvalue", "ccrm_disciplinesvalue", "ccrm_projectsectorvalue");
            query.Criteria = new FilterExpression();
            query.Criteria.FilterOperator = LogicalOperator.Or;
            query.Criteria.AddCondition("ccrm_othernetworksval", ConditionOperator.NotNull);
            query.Criteria.AddCondition("ccrm_servicesvalue", ConditionOperator.NotNull);
            query.Criteria.AddCondition("ccrm_theworksvalue", ConditionOperator.NotNull);
            query.Criteria.AddCondition("ccrm_disciplinesvalue", ConditionOperator.NotNull);
            query.Criteria.AddCondition("ccrm_projectsectorvalue", ConditionOperator.NotNull);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            EntityCollection final = new EntityCollection();
            foreach (Entity i in entityCollection.Entities)
            {
                final.Entities.Add(i);
                UpdateOpportunityMultiSelect(service,
                    i.GetAttributeValue<Guid>("opportunityid"),
                    i.GetAttributeValue<string>("ccrm_othernetworksval"),
                    i.GetAttributeValue<string>("ccrm_servicesvalue"),
                    i.GetAttributeValue<string>("ccrm_theworksvalue"),
                    i.GetAttributeValue<string>("ccrm_disciplinesvalue"),
                    i.GetAttributeValue<string>("ccrm_projectsectorvalue"));
            }
            do
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                entityCollection = service.RetrieveMultiple(query);
                foreach (Entity i in entityCollection.Entities)
                {
                    final.Entities.Add(i);
                    UpdateOpportunityMultiSelect(service,
                    i.GetAttributeValue<Guid>("opportunityid"),
                    i.GetAttributeValue<string>("ccrm_othernetworksval"),
                    i.GetAttributeValue<string>("ccrm_servicesvalue"),
                    i.GetAttributeValue<string>("ccrm_theworksvalue"),
                    i.GetAttributeValue<string>("ccrm_disciplinesvalue"),
                    i.GetAttributeValue<string>("ccrm_projectsectorvalue"));
                }
            }
            while (entityCollection.MoreRecords);
            Console.WriteLine("Total Opportunity record count:" + final.TotalRecordCount);
            Console.ReadKey();
            string str = "";
        }

        //ccrm_othernetworksval", "ccrm_servicesvalue", "ccrm_theworksvalue", "ccrm_disciplinesvalue", "ccrm_projectsectorvalue"
        public static void UpdateOpportunityMultiSelect(IOrganizationService service, Guid opportunityId, string othernetworksval, string servicesvalue, string theworksvalue, string disciplinesvalue, string projectsectorvalue)
        {
            Entity opportunity = new Entity("opportunity");
            if (othernetworksval != string.Empty && othernetworksval != null)
            {
                OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                string[] arr = othernetworksval.Split(',');
                foreach (var item in arr)
                {
                    collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                }

                opportunity["arup_globalservices"] = collectionOptionSetValues;
            }
            if (servicesvalue != string.Empty && servicesvalue != null)
            {
                OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                string[] arr = servicesvalue.Split(',');
                foreach (var item in arr)
                {
                    collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                }

                opportunity["arup_services"] = collectionOptionSetValues;
            }
            if (theworksvalue != string.Empty && theworksvalue != null)
            {
                OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                string[] arr = theworksvalue.Split(',');
                foreach (var item in arr)
                {
                    collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                }

                opportunity["arup_projecttype"] = collectionOptionSetValues;
            }
            if (disciplinesvalue != string.Empty && disciplinesvalue != null)
            {
                OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                string[] arr = disciplinesvalue.Split(',');
                foreach (var item in arr)
                {
                    collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                }

                opportunity["arup_disciplines"] = collectionOptionSetValues;
            }
            if (projectsectorvalue != string.Empty && projectsectorvalue != null)
            {
                OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                string[] arr = projectsectorvalue.Split(',');
                foreach (var item in arr)
                {
                    collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                }

                opportunity["arup_projectsector"] = collectionOptionSetValues;
            }
            opportunity.Id = opportunityId;
            service.Update(opportunity);
        }
        #endregion
    }
}

