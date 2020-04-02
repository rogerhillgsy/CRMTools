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

namespace ModernOdataApp.Models
{
    public class TasksModel : INotifyPropertyChanged
    {
        private string _subject;
        /// <summary>
        /// Tasks ViewModel property; this property is used in the view to display its value using a binding.
        /// </summary>
        /// <returns></returns>
        public string Subject
        {
            get
            {
                return _subject;
            }
            set
            {
                if (value != _subject)
                {
                    _subject = value;
                    NotifyPropertyChanged("Subject");
                }
            }
        }

        private DateTime _scheduledStartDate;
        /// <summary>
        /// Tasks ViewModel property; this property is used in the view to display its value using a binding.
        /// </summary>
        /// <returns></returns>
        public DateTime ScheduledStartDate
        {
            get
            {
                return _scheduledStartDate;
            }
            set
            {
                if (value != _scheduledStartDate)
                {
                    _scheduledStartDate = value;
                    NotifyPropertyChanged("ScheduledStartDate");
                }
            }
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
