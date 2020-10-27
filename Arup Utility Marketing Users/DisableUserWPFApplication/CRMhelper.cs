using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace DisableUserWPFApplication
{
    class CRMHelper
    {
        public IOrganizationService service;

        public CRMHelper(string UserName, string Password, string Url)
        {
            //ClientCredentials cre = new ClientCredentials();

            //cre.UserName.UserName = "Global\\" + UserName;

            //cre.UserName.Password = Password;

            //using (OrganizationServiceProxy proxy = new OrganizationServiceProxy(new Uri(Url + "/XRMServices/2011/Organization.svc"), null, cre, null))

            //{

            //    proxy.EnableProxyTypes();

            //    service = (IOrganizationService)proxy;

            //}

            string connectionstring = "Url="+ Url+ ";Username=" + UserName +";Password="+Password+";authtype=Office365;";
            var CrmService = new CrmServiceClient(connectionstring);
            service = CrmService.OrganizationServiceProxy;

        }
        public int EnableUserIfDisabled(User user)
        {
            try
            {
                if (user.status == true)
                {

                    Entity userEntity = service.Retrieve("systemuser", user.UserId, new ColumnSet(new string[] { "systemuserid", "isdisabled" }));
                    SetStateRequest request = new SetStateRequest()
                    {
                        EntityMoniker = userEntity.ToEntityReference(),
                        State = new OptionSetValue(0),
                        Status = new OptionSetValue(-1)
                    };
                    service.Execute(request);
                    return 0;

                }
                else
                {
                    //Message += "User already enabled.";
                    return 1;
                }
            }
            catch (Exception ex)
            {
                //  Message += ex.Message + ".";
                throw ex;

            }
        }

        public int DisableUserIfEnabled(User user)
        {
            try
            {
                if (user.status == false)
                {

                    Entity userEntity = service.Retrieve("systemuser", user.UserId, new ColumnSet(new string[] { "systemuserid", "isdisabled" }));
                    SetStateRequest request = new SetStateRequest()
                    {
                        EntityMoniker = userEntity.ToEntityReference(),
                        State = new OptionSetValue(1),
                        Status = new OptionSetValue(-1)
                    };
                    service.Execute(request);
                    return 0;

                }
                else
                {
                    //Message += "User already disabled.";
                    return 1;
                }
            }
            catch (Exception ex)
            {
                //  Message += ex.Message + ".";
                throw ex;

            }
        }
        public User getUserIdFromCRM(string domainUserName)
        {
            User returnUser = new User();
            string fetchXMLUser = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
              <entity name='systemuser'>
                <attribute name='fullname' />
                <attribute name='systemuserid' />
                <attribute name='caltype' />
                <attribute name='isdisabled' />
                <order attribute='fullname' descending='false' />
                <filter type='and'>
                  <condition attribute='domainname' operator='eq' value='" + domainUserName + @"' />
                  <condition attribute='fullname' operator='ne' value='SYSTEM' />
                  <condition attribute='fullname' operator='ne' value='INTEGRATION' />
                </filter>
              </entity>
            </fetch>";
            try
            {

                EntityCollection usersFetched = service.RetrieveMultiple(new FetchExpression(fetchXMLUser));
                if (usersFetched.Entities.Count == 0)
                {
                    throw new Exception("No User found with Staff-Id:" + domainUserName);

                }
                else if (usersFetched.Entities.Count > 1)
                {
                    throw new Exception("Multiple Users found with Staff-Id:" + domainUserName);
                }
                else if (usersFetched.Entities.Count == 1)
                {
                    returnUser.UserId = usersFetched[0].Id;
                    returnUser.FullName = usersFetched[0]["fullname"].ToString();
                    returnUser.LicenseType = ((OptionSetValue)(usersFetched[0]["caltype"])).Value;
                    returnUser.status = Convert.ToBoolean(usersFetched[0]["isdisabled"]);
                    returnUser.DomainUserName = domainUserName;
                }
                else
                {
                    throw new Exception("Some other error occurred:" + domainUserName);

                }

            }
            catch (Exception ex)
            {
                //Message += ex.Message + ".";
                throw ex;
            }


            return returnUser;
        }

        public EntityCollection ExecuteFetch(string fetchxml)
        {
            EntityCollection fetchResults = service.RetrieveMultiple(new FetchExpression(fetchxml));
            return fetchResults;
        }

        public void AddlistMembers(string marketingListGUID,Guid memberguid)
        {
            try
            {
                //AddMemberListRequest req = new AddMemberListRequest();
                ////we will add an account to our marketing list
                ////entity type must be an account, contact, or lead
                //req.EntityId = memberguid;
                
                ////we will add the account to this existing marketing list
                //req.ListId = new Guid(marketingListGUID);
                //AddMemberListResponse resp = (AddMemberListResponse)service.Execute(req);

                var addMemberListReq = new AddListMembersListRequest
                {
                    MemberIds = new[] { memberguid },
                    ListId = Guid.Parse(marketingListGUID)
                };
                AddListMembersListResponse resp = (AddListMembersListResponse)service.Execute(addMemberListReq);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        public EntityCollection RetrieveDynamicMemberList(string strList)
        {
            QueryByAttribute query = new QueryByAttribute("listmember");
            // pass the guid of the Static marketing list
            query.AddAttributeValue("listid", new Guid(strList));
            query.ColumnSet = new ColumnSet(true);
            EntityCollection entityCollection = service.RetrieveMultiple(query);
            return entityCollection;
        }
    }

   
}
