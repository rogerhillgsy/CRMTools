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
using ModernOdataApp.Models;
using ModernOdataApp;
using System.Xml.Linq;


namespace Sample.ViewModels
{
    public class TasksViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<TasksModel> _tasks;
        public ObservableCollection<TasksModel> Tasks
        {
            get { return _tasks; }
            set
            {
                if (value != _tasks)
                {
                    _tasks = value;
                    NotifyPropertyChanged("Tasks");
                }
            }
        }

        /// <summary>
        /// Fetch Tasks details.
        /// Extracts Tasks details from the XML response and binds data to observable collection.
        /// </summary>  
        public async Task<ObservableCollection<TasksModel>> LoadTasksData(string AccessToken)
        {
            var TasksResponseBody = await HttpRequestBuilder.Retrieve(AccessToken, new string[] { "Subject", "ScheduledStart" }, "TaskSet");
            Tasks = new ObservableCollection<TasksModel>();
          
            // Feed namespace
            XNamespace rss = "http://www.w3.org/2005/Atom"; 
            // d namespace
            XNamespace d = "http://schemas.microsoft.com/ado/2007/08/dataservices";
            // m namespace
            XNamespace m = "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata";

            // Convert the string response to an xDocument.
            XDocument xdoc = XDocument.Parse(TasksResponseBody.ToString(), LoadOptions.None);
            foreach (var entry in xdoc.Root.Descendants(rss + "entry").Descendants(rss + "content").Descendants(m + "properties"))
            {
                TasksModel task=new TasksModel();                
                task.Subject= entry.Element(d + "Subject").Value;
                task.ScheduledStartDate= DateTime.Parse(entry.Element(d + "ScheduledStart").Value);
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
