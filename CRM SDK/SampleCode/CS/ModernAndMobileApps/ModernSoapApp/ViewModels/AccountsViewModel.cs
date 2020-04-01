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

using ModernSoapApp;
using ModernSoapApp.Models;
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
        private ObservableCollection<AccountsModel> accounts;
        public ObservableCollection<AccountsModel> Accounts
        {
            get { return accounts; }
            set
            {
                if (value != accounts)
                {
                    accounts = value;
                    NotifyPropertyChanged("Accounts");
                }
            }
        }

        /// <summary>
        /// Fetch Accounts details.
        /// Extracts Accounts details from XML response and binds data to Observable Collection.
        /// </summary>    
        public async Task<ObservableCollection<AccountsModel>> LoadAccountsData(string AccessToken)
        {
            var AccountsResponseBody = await HttpRequestBuilder.RetrieveMultiple(AccessToken, new string[] { "name", "emailaddress1", "telephone1" }, "account");

            Accounts = new ObservableCollection<AccountsModel>();

            // Converting response string to xDocument.
            XDocument xdoc = XDocument.Parse(AccountsResponseBody.ToString(), LoadOptions.None);
            XNamespace s = "http://schemas.xmlsoap.org/soap/envelope/";//Envelop namespace s
            XNamespace a = "http://schemas.microsoft.com/xrm/2011/Contracts";//a namespace
            XNamespace b = "http://schemas.datacontract.org/2004/07/System.Collections.Generic";//b namespace

            foreach (var entity in xdoc.Descendants(s + "Body").Descendants(a + "Entities").Descendants(a + "Entity"))
            {
                AccountsModel account = new AccountsModel();
                foreach (var KeyvaluePair in entity.Descendants(a + "KeyValuePairOfstringanyType"))
                {
                    if (KeyvaluePair.Element(b + "key").Value == "name")
                    {
                        account.Name = KeyvaluePair.Element(b + "value").Value;
                    }
                    else if (KeyvaluePair.Element(b + "key").Value == "emailaddress1")
                    {
                        account.Email = KeyvaluePair.Element(b + "value").Value;
                    }
                    else if (KeyvaluePair.Element(b + "key").Value == "telephone1")
                    {
                        account.Phone = KeyvaluePair.Element(b + "value").Value;
                    }
                }
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

