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

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ModernSoapApp.Models;
using ModernSoapApp;
using System.Xml.Linq;

namespace Sample.ViewModels
{
    public class TasksViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<TasksModel> tasks;
        public ObservableCollection<TasksModel> Tasks
        {
            get { return tasks; }
            set
            {
                if (value != tasks)
                {
                    tasks = value;
                    NotifyPropertyChanged("Tasks");
                }
            }
        }

        /// <summary>
        /// Fetch Tasks details.
        /// Extracts Tasks details from XML response and binds data to Observable Collection.
        /// </summary>  
        public async Task<ObservableCollection<TasksModel>> LoadTasksData(string AccessToken)
        {
            var TasksResponseBody = await HttpRequestBuilder.RetrieveMultiple(AccessToken, new string[] { "subject", "scheduledstart" }, "task");

            Tasks = new ObservableCollection<TasksModel>();

            // Converting response string to xDocument.
            XDocument xdoc = XDocument.Parse(TasksResponseBody.ToString(), LoadOptions.None);
            XNamespace s = "http://schemas.xmlsoap.org/soap/envelope/";//Envelop namespace s
            XNamespace a = "http://schemas.microsoft.com/xrm/2011/Contracts";//a namespace
            XNamespace b = "http://schemas.datacontract.org/2004/07/System.Collections.Generic";//b namespace

            foreach (var entity in xdoc.Descendants(s + "Body").Descendants(a + "Entities").Descendants(a + "Entity"))
            {
                TasksModel task = new TasksModel();
                foreach (var KeyvaluePair in entity.Descendants(a + "KeyValuePairOfstringanyType"))
                {                   
                    if (KeyvaluePair.Element(b + "key").Value == "subject")
                    {
                        task.Subject = (string)KeyvaluePair.Element(b + "value").Value;
                    }
                    else if (KeyvaluePair.Element(b + "key").Value == "scheduledstart")
                    {
                        task.ScheduledStartDate = DateTime.Parse(KeyvaluePair.Element(b + "value").Value);
                    }

                }
                Tasks.Add(task);
            }
            return Tasks;
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
