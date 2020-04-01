// =====================================================================
//  This file is part of the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================

using ModernOdataApp;
using ModernOdataApp.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Sample.ViewModels
{
    public class AccountsViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<AccountsModel> _accounts;
        public ObservableCollection<AccountsModel> Accounts
        {
            get { return _accounts; }
            set
            {
                if (value != _accounts)
                {
                    _accounts = value;
                    NotifyPropertyChanged("Accounts");
                }
            }
        }

        /// <summary>
        /// Fetch accounts details.
        /// Extracts accounts details from the XML response and bind the data to an observable collection.
        /// </summary>    
        public async Task<ObservableCollection<AccountsModel>> LoadAccountsData(string AccessToken)
        {
            var AccountsResponseBody = await HttpRequestBuilder.Retrieve(AccessToken, new string[] { "Name", "EMailAddress1", "Telephone1" }, "AccountSet");

            Accounts = new ObservableCollection<AccountsModel>();

            //feed namespace
            XNamespace rss = "http://www.w3.org/2005/Atom";
            // d namespace
            XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";
            // m namespace
            XNamespace m = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";

            // Convert the string response to an xDocument.
            XDocument xdoc = XDocument.Parse(AccountsResponseBody.ToString(), LoadOptions.None);

            foreach (var entry in xdoc.Root.Descendants(rss + "entry").Descendants(rss + "content").Descendants(m + "properties"))
            {
                AccountsModel account=new AccountsModel();
                account.Name= entry.Element(d + "Name").Value ;
                account.Email= entry.Element(d + "EMailAddress1").Value;
                account.Phone= entry.Element(d + "Telephone1").Value;   
                Accounts.Add(account);
            }                            
            return Accounts;
        }

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        // Used to notify Silverlight that a property has changed.
        private void NotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion
    }
}
