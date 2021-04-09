using DisableUserWPFApplication.Models;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DisableUserWPFApplication
{
    public class MarketingBL
    {
        private string userName = "";
        private string password = "";
        private string url = "";

        public MarketingBL(string userName,string password, string url)
        {
            this.url = url;
            this.userName = userName;
            this.password = password;
        }

      
        public List<MarketingListModel> getMarketingList()
        {
            List<MarketingListModel> list = new List<MarketingListModel>();
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='list'>
                                    <attribute name='listname' />
                                    <attribute name='listid' />
                                    <order attribute='listname' descending='false' />
                                     <filter type='and'>
                                        <condition attribute = 'statecode' operator= 'eq' value = '0' />
                                        <condition attribute = 'type' operator= 'eq' value = '0' />
                                        <condition attribute='createdfromcode' operator='eq' value='2' />
                                         </filter>
                                     </entity>
                                </fetch>";

            CRMHelper _crmHelper = new CRMHelper(userName, password, url);
            var collection = _crmHelper.ExecuteFetch(fetchXML);

            if (collection.Entities.Count >= 0)
            {
                int i = 1;
                foreach (var item in collection.Entities)
                { 
                    MarketingListModel mlist = new MarketingListModel();
                    mlist.ID = i;
                    mlist.GuidValue = item["listid"].ToString();
                    mlist.Name = item["listname"].ToString();
                    list.Add(mlist);
                    i++;
                }
            }
            else 
            {
                throw new Exception("Marketing list not found ");
            }
            return list;
        }

       

        public List<ContactModel> getContactDetails(string email)
        {
            List<ContactModel> lstContacts = new List<ContactModel>();
            string fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='contact'>
                                    <attribute name='fullname' />
                                    <attribute name='contactid' />
                                    <attribute name='statuscode' />
                                    <order attribute='fullname' descending='false' />
                                    <filter type='and'>
                                  <condition attribute='emailaddress1' operator='eq' value= '" + email + "' />  </filter> </entity> </fetch>";


            CRMHelper _crmHelper = new CRMHelper(userName, password, url);
            var collection = _crmHelper.ExecuteFetch(fetchxml);

            if (collection.Entities.Count >= 0)
            {
                foreach (var item in collection.Entities)
                {
                    ContactModel contact = new ContactModel();
                    contact.ID = Guid.Parse(item["contactid"].ToString());
                    contact.Name = item["fullname"].ToString();
                    contact.status = ((Microsoft.Xrm.Sdk.OptionSetValue)(item["statuscode"])).Value.ToString();
                    lstContacts.Add(contact);
                }
            }

            return lstContacts;
        }

        public List<MemberDetails> getMarketingListMembers(string guidValue)
        {
            CRMHelper _crmHelper = new CRMHelper(userName, password, url);
            List<MemberDetails> members = new List<MemberDetails>();
            var collection = _crmHelper.RetrieveDynamicMemberList(guidValue);
            if (collection != null)
            {
                foreach (var items in collection.Entities)
                {
                  
                    if (items.Attributes.Contains("entityid"))
                    {
                        MemberDetails objMember = new MemberDetails();
                        objMember.Memberid = ((Microsoft.Xrm.Sdk.EntityReference)(items.Attributes["entityid"])).Id;
                        objMember.Name = ((Microsoft.Xrm.Sdk.EntityReference)(items.Attributes["entityid"])).Name;
                        members.Add(objMember);
                    }

                }
            }
            return members;
        }
        public void AddMembersToList(string marketingListGUID, Guid memberguids)
        {
            try
            {
                CRMHelper _crmHelper = new CRMHelper(userName, password, url);
                _crmHelper.AddlistMembers(marketingListGUID, memberguids);

            }
            catch(Exception ex)
            {
                throw ex;
            }
            

        }
    }
}
