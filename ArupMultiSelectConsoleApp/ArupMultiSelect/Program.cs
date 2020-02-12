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

namespace ArupMultiSelect
{
    public class Program
    {
        static List<string> linesInFailedFile = null;
        public static void Main(string[] args)
        {
            try
            {
                string serverUrl = ConfigurationManager.AppSettings["serverUrl"].ToString();
                string userName = ConfigurationManager.AppSettings["UserName"].ToString();
                string password = ConfigurationManager.AppSettings["Password"].ToString();
                string domain = ConfigurationManager.AppSettings["Password"].ToString();
                IOrganizationService service = CreateService1(serverUrl, userName, password, domain);
                //IOrganizationService service = CreateService1("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc", "crm.hub@arup.com", "CIm2$98pRt", "arup");
                UpdateContact(service);
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


        public static IOrganizationService CreateService1(string serverUrl, string userId, string password, string domain)
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

        public static IOrganizationService CreateService(string serverUrl, string userId, string password, string domain)
        {
            IOrganizationService organizationService = null;

            try
            {

                ClientCredentials clientCredentials = new ClientCredentials();
                clientCredentials.UserName.UserName = "sreenath.medikurthi@arup.com";
                clientCredentials.UserName.Password = "KYX76t5HV8b0";

                // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                // Get the URL from CRM, Navigate to Settings -> Customizations -> Developer Resources
                // Copy and Paste Organization Service Endpoint Address URL
                organizationService = (IOrganizationService)new OrganizationServiceProxy(new Uri("https://arupgroupcloud.crm4.dynamics.com/XRMServices/2011/Organization.svc"),
                 null, clientCredentials, null);
                //organizationService = (IOrganizationService)new OrganizationServiceProxy(new Uri("https://dcsguardiandevint0.crm9.dynamics.com/XRMServices/2011/Organization.svc"),
                // null, clientCredentials, null);
                if (organizationService != null)
                {
                    Guid userid = ((WhoAmIResponse)organizationService.Execute(new WhoAmIRequest())).UserId;

                    if (userid != Guid.Empty)
                    {
                        Console.WriteLine("Connection Established Successfully...");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to Established Connection!!!");
                }
                return organizationService;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught - " + ex.Message);
                return organizationService;
            }
        }

        public static IOrganizationService CreateService2(string serverUrl, string userId, string password, string domain)
        {
            //objDataValidation.CreateLog("Before CRm  Creation");
            IOrganizationService _service;

            ClientCredentials Credentials = new ClientCredentials();
            ClientCredentials devivceCredentials = new ClientCredentials();

            Credentials.UserName.UserName = domain + "\\" + userId;
            //Credentials.UserName.UserName = userId;
            Credentials.UserName.Password = password;
            //Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
            //This URL needs to be updated to match the servername and Organization for the environment.
            //The following URLs should be used to access the Organization service(SOAP endpoint):
            //https://{Organization Name}.api.crm.dynamics.com/XrmServices/2011/Organization.svc (North America)
            //https://{Organization Name}.api.crm4.dynamics.com/XrmServices/2011/Organization.svc (EMEA)
            //https://{Organization Name}.api.crm5.dynamics.com/XrmServices/2011/Organization.svc (APAC)
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
                return _service;
            }
            catch (Exception ex)
            {
                throw new System.Exception("<Error>Problem in creating CRM Service</Error>" + ex.Message);
            }
        }
        #endregion
             
        public static void UpdateContact(IOrganizationService service)
        {
            
            QueryExpression query = new QueryExpression("contact");
            //query.ColumnSet = new ColumnSet("ccrm_businessinterestpicklistname", "ccrm_businessinterestpicklistvalue", "arup_businessinterest");
            query.ColumnSet.AddColumns("ccrm_businessinterestpicklistname", "ccrm_businessinterestpicklistvalue", "arup_businessinterest");
            query.Criteria = new FilterExpression();

            query.Criteria.AddCondition("ccrm_businessinterestpicklistvalue", ConditionOperator.NotNull);
            query.PageInfo = new PagingInfo();
            query.PageInfo.Count = 5;
            query.PageInfo.PageNumber = 1;
            query.PageInfo.ReturnTotalRecordCount = true;
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            EntityCollection final = new EntityCollection();
            foreach (Entity i in entityCollection.Entities)
            {
                final.Entities.Add(i);
                UpdateContactMultiSelect(service, i.GetAttributeValue<Guid>("contactid"),i.GetAttributeValue<string>("ccrm_businessinterestpicklistvalue"));
            }
            do
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                entityCollection = service.RetrieveMultiple(query);
                foreach (Entity i in entityCollection.Entities)
                {
                    final.Entities.Add(i);
                    UpdateContactMultiSelect(service, i.GetAttributeValue<Guid>("contactid"), i.GetAttributeValue<string>("ccrm_businessinterestpicklistvalue"));
                }
            }
            while (entityCollection.MoreRecords);
            Console.WriteLine("Total Contact record count:"+ final.TotalRecordCount);
            Console.ReadKey();
            string str = "";
        }

        public static void UpdateContactMultiSelect(IOrganizationService service, Guid contactId, string businessInterestValues)
        {
            try
            {
                if (businessInterestValues != string.Empty && businessInterestValues != null)
                {
                    Entity contact = new Entity("contact");
                    OptionSetValueCollection collectionOptionSetValues = new OptionSetValueCollection();
                    string[] arr = businessInterestValues.Split(',');

                    foreach (var item in arr)
                    {
                        collectionOptionSetValues.Add(new OptionSetValue(Convert.ToInt32(item)));
                    }

                    contact["arup_businessinterest"] = collectionOptionSetValues;
                    contact.Id = contactId;
                    service.Update(contact);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error : " + e.Message);
                //linesInFailedFile.Add("RecordId,ToOptionset,Values,Message");
                string optionSetValues = "arup_businessinterest : " + businessInterestValues;

                linesInFailedFile.Add(string.Format("{0},{1},{2},{3}", "Contact", contactId, e.Message, optionSetValues));
            }
        }
    }
}
