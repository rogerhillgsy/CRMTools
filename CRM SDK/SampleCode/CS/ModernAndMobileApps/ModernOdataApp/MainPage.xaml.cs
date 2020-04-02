// =====================================================================
//
//  This file is part of the Microsoft Dynamics CRM SDK Code Samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or online documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
//
// =====================================================================

using System;
using ModernOdataApp.Models;
using ModernOdataApp.Views;
using Windows.UI.Core;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;
using Windows.Storage;
using System.Collections.ObjectModel;
using Windows.UI.Popups;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using ModernOdataApp.Common;
using Windows.Security.Authentication.Web;
// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace ModernOdataApp
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        #region Class Level Members

        private string _accessToken = string.Empty;
        private static string _strItemClicked;                     
        // Currently using static fields as menu items for displaying
        private string[] _strMenuItems = new string[] { "Leads", "Opportunities", "Accounts", "Contacts", "Tasks", "Placeholder", "Placeholder" };
        private ObservableCollection<MainPageItem> _theMenuItems { get; set; }

        # endregion

        public MainPage()
        {
            this.InitializeComponent();
        }

        /// <summary>
        /// Invoked when this page is about to be displayed in a Frame.
        /// </summary>
        /// <param name="e">Event data that describes how this page was reached.  The Parameter
        /// property is typically used to configure the page.</param>
        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            progressBar.Visibility = Visibility.Visible;        
            Initialize();
        }

        /// <summary>
        /// Binding Menu items to Main Page
        /// </summary>
        private async void Initialize()
        {
            _accessToken = await CurrentEnvironment.Initialize();

            pageTitle.Text = "Welcome to the Windows 8 Modern OData App";
            _theMenuItems = new ObservableCollection<MainPageItem>();
            for (int i = 0; i < 7; i++)
            {
                MainPageItem anItem = new MainPageItem()
                {
                    Name = _strMenuItems[i]
                };
                _theMenuItems.Add(anItem);
            }
            itemsViewSource.Source = _theMenuItems;
            progressBar.Visibility = Visibility.Collapsed;
        }
        
        private async void NavigateTo(Type pageType, object parameter)
        {
            await Dispatcher.RunAsync(CoreDispatcherPriority.Normal, () => Frame.Navigate(pageType, parameter));
        }

        private void itemGridView_ItemClick(object sender, ItemClickEventArgs e)
        {
            MainPageItem selItem = (MainPageItem)e.ClickedItem;
            _strItemClicked = selItem.Name;
            if (_strItemClicked.Equals("Tasks"))
            {
                this.NavigateTo(typeof(Tasks), _accessToken);
            }

            else if (_strItemClicked.Equals("Accounts"))
            {
                this.NavigateTo(typeof(Accounts), _accessToken);
            }
        }

    }
}
