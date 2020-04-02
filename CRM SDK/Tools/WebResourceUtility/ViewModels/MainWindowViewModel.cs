using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Xml.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using WebResourceUtility.Model;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Windows.Threading;

namespace Microsoft.Crm.Sdk.Samples
{
    public class MainWindowViewModel : ViewModelBase
    {

        #region Commands


        RelayCommand _hideOutputWindow;
        public ICommand HideOutputWindow 
        {
            get
            {
                if (_hideOutputWindow == null)
                {
                    _hideOutputWindow = new RelayCommand(
                        param => { IsOutputWindowDisplayed = false; });
                }
                return _hideOutputWindow;
            }
        }

        RelayCommand _showOutputWindow;
        public ICommand ShowOutputWindow
        {
            get
            {
                if (_showOutputWindow == null)
                {
                    _showOutputWindow = new RelayCommand(
                        param => { IsOutputWindowDisplayed = true; });
                }
                return _showOutputWindow;
            }
        }

        RelayCommand _browseFolderCommand;
        public ICommand BrowseFolderCommand 
        {
            get
            {
                if (_browseFolderCommand == null)
                {
                    _browseFolderCommand = new RelayCommand(
                        param => this.ShowBrowseFolderDialog());
                }
                return _browseFolderCommand;
            }

        }

        RelayCommand _activateConnectionCommand;
        public ICommand ActivateConnectionCommand 
        {
            get
            {
                if (_activateConnectionCommand == null)
                {
                    _activateConnectionCommand = new RelayCommand(
                        param => this.ActivateSelectedConfiguration(),
                        param => this.CanActivateSelectedConfiguration());
                }
                return _activateConnectionCommand;
            }

        }

        RelayCommand _createNewConnectionCommand;
        public ICommand CreateNewConnectionCommand
        {
            get
            {
                if (_createNewConnectionCommand == null)
                {
                    _createNewConnectionCommand = new RelayCommand(
                        param => this.CreateNewConfiguration());
                }
                return _createNewConnectionCommand;
            }

        }
                
        RelayCommand _deleteConnectionCommand;
        public ICommand DeleteConnectionCommand
        {
            get
            {
                if (_deleteConnectionCommand == null)
                {
                    _deleteConnectionCommand = new RelayCommand(
                        param => this.DeleteSelectedConfiguration());
                }
                return _deleteConnectionCommand;
            }

        }

        RelayCommand _activateSolutionCommand;
        public ICommand ActivateSolutionCommand
        {
            get
            {
                if (_activateSolutionCommand == null)
                {
                    _activateSolutionCommand = new RelayCommand(
                        param => this.ActivateSelectedSolution());
                }
                return _activateSolutionCommand;
            }
        }

        RelayCommand _activateSelectedPackageCommand;
        public ICommand ActivateSelectedPackageCommand
        {
            get
            {
                if (_activateSelectedPackageCommand == null)
                {
                    _activateSelectedPackageCommand = new RelayCommand(
                        param => this.ActivatePackage());
                }
                return _activateSelectedPackageCommand;
            }
        }

        RelayCommand _createNewPackageCommand;
        public ICommand CreateNewPackageCommand
        {
            get
            {
                if (_createNewPackageCommand == null)
                {
                    _createNewPackageCommand = new RelayCommand(
                        param => this.CreateNewPackage());
                }
                return _createNewPackageCommand;
            }
        }

        RelayCommand _deleteActivePackageCommand;
        public ICommand DeleteActivePackageCommand
        {
            get
            {
                if (_deleteActivePackageCommand == null)
                {
                    _deleteActivePackageCommand = new RelayCommand(
                        param => this.DeleteSelectedPackage());
                }
                return _deleteActivePackageCommand;
            }

        }

        RelayCommand _saveActivePackageCommand;
        public ICommand SaveActivePackageCommand
        {
            get
            {
                if (_saveActivePackageCommand == null)
                {
                    _saveActivePackageCommand = new RelayCommand(
                        param => SavePackages());
                }
                return _saveActivePackageCommand;
            }
        }            

        RelayCommand _refreshFilesCommand;
        public ICommand RefreshFilesCommand
        {
            get
            {
                if (_refreshFilesCommand == null)
                {
                    _refreshFilesCommand = new RelayCommand(
                        param => this.SearchAndPopulateFiles());
                }
                return _refreshFilesCommand;
            }
        }

        RelayCommand _saveConnectionsCommand;
        public ICommand SaveConnectionsCommand
        {
            get
            {
                if (_saveConnectionsCommand == null)
                {
                    _saveConnectionsCommand = new RelayCommand(
                        param => SaveConfigurations());
                }
                return _saveConnectionsCommand;
            }
        }

        RelayCommand<IEnumerable> _convertFileToResourceCommand;
        public ICommand ConvertFileToResourceCommand
        {
            get
            {
                if (_convertFileToResourceCommand == null)
                {
                    _convertFileToResourceCommand = new RelayCommand<IEnumerable>(AddFilesToWebResources);
                }
                return _convertFileToResourceCommand;
            }
        }

        RelayCommand<IEnumerable> _uploadWebResourcesCommand;
        public ICommand UploadWebResourcesCommand
        {
            get
            {
                if (_uploadWebResourcesCommand == null)
                {
                    _uploadWebResourcesCommand = new RelayCommand<IEnumerable>(UploadWebResources, param => CanUploadWebResource());
                }
                return _uploadWebResourcesCommand;
            }
        }

        RelayCommand _uploadAllWebResourcesCommand;
        public ICommand UploadAllWebResourcesCommand
        {
            get
            {
                if (_uploadAllWebResourcesCommand == null)
                {
                    _uploadAllWebResourcesCommand = 
                        new RelayCommand(param =>                    UploadAllWebResources(), 
                            param => CanUploadWebResource());
                }
                return _uploadAllWebResourcesCommand;
            }
        }

        RelayCommand<IEnumerable> _deleteWebResourcesCommand;
        public ICommand DeleteWebResourcesCommand
        {
            get
            {
                if (_deleteWebResourcesCommand == null)
                {
                    _deleteWebResourcesCommand = new RelayCommand<IEnumerable>(DeleteSelectedWebResources);
                }
                return _deleteWebResourcesCommand;
            }
        }
        
        #endregion 
        
        #region Properties

        public const string CONFIG_FILENAME = @"configurations.xml";
        public const string PACKAGES_FILENAME = @"packages.xml";
        public const string VALID_NAME_MSG = "ERROR: Web Resource names cannot contain spaces or hyphens. They must be alphanumeric and contain underscore characters, periods, and non-consecutive forward slash characters";
        public XElement XmlPackageData;
        public XElement XmlConfigData;

        private StringBuilder _progressMessage;
        public String ProgressMessage
        {
            get
            {
                return _progressMessage.ToString();
            }
            set
            {
                _progressMessage.AppendLine(value);
                OnPropertyChanged("ProgressMessage");
            }
        }

        private int _tabControlSelectedIndex;
        public int TabControlSelectedIndex
        {
            get
            {
                return _tabControlSelectedIndex;
            }
            set
            {
                _tabControlSelectedIndex = value;
                OnPropertyChanged("TabControlSelectedIndex");
            }
        }

        private bool _areAllButtonsEnabled = true;
        public bool AreAllButtonsEnabled
        {
            get
            {
                return _areAllButtonsEnabled;
            }
            set
            {
                _areAllButtonsEnabled = value;
                OnPropertyChanged("AreAllButtonsEnabled");
            }
        }

        public bool IsActiveConnectionSet
        {
            get 
            {
                return (ActiveConfiguration != null) ? true : false;
            }
        }
        public bool IsActiveSolutionSet
        {
            get
            {
                return (ActiveSolution != null) ? true : false;
            }
        }
        public bool IsActivePackageSet
        {
            get
            {
                return (ActivePackage != null) ? true : false;
            }
        }

        private bool _shouldPublishAllAfterUpload;
        public bool ShouldPublishAllAfterUpload
        {
            get
            {
                return _shouldPublishAllAfterUpload;
            }
            set 
            {
                _shouldPublishAllAfterUpload = value;
                OnPropertyChanged("ShouldPublishAllAfterUpload");
            }
        }

        private bool _isOutputWindowDisplayed = false;
        public bool IsOutputWindowDisplayed
        {
            get
            {
                return _isOutputWindowDisplayed;
            }
            set
            {
                _isOutputWindowDisplayed = value;
                OnPropertyChanged("IsOutputWindowDisplayed");
                OnPropertyChanged("IsWorkstationDisplayed");
            }
        }

        public bool IsWorkstationDisplayed
        {
            get
            {
                return !(IsOutputWindowDisplayed);
            }
        }

        private String _fileSearchText;
        public String FileSearchText
        {
            get { return _fileSearchText; }
            set { _fileSearchText = value; OnPropertyChanged("FileSearchText"); }
        }
        
        //WebResources Packages
        public ObservableCollection<XElement> Packages { get; set; }
        private XElement _selectedPackage;
        public XElement SelectedPackage
        {
            get
            {
                return _selectedPackage;
            }
            set
            {
                _selectedPackage = value;
                OnPropertyChanged("SelectedPackage");
            }
        }
        private XElement _activePackage;
        public XElement ActivePackage
        {
            get
            {
                return _activePackage;
            }
            set
            {
                _activePackage = value;
                OnPropertyChanged("ActivePackage");
                OnPropertyChanged("IsActivePackageSet");
            }
        }
        private bool _isActivePackageDirty = false;
        public bool IsActivePackageDirty
        {
            get
            {
                return _isActivePackageDirty;
            }
            set
            {
                _isActivePackageDirty = value;
                OnPropertyChanged("IsActivePackageDirty");
            }
        }

        //FileInfos for all potential resources in a directory
        public ObservableCollection<FileInfo> CurrentFiles { get; set; }
        public ObservableCollection<FileInfo> CurrentFilesSelected { get; set; }

        //Represents a collection of "WebResourceInfo" node from XML.
        public ObservableCollection<XElement> WebResourceInfos { get; set; }
        public ObservableCollection<XElement> WebResourceInfosSelected { get; set; }
        
        //Connections
        public ObservableCollection<XElement> Configurations { get; set; }
        private XElement _selectedConfiguration;
        public XElement SelectedConfiguration 
        {
            get { return _selectedConfiguration; }
            set
            {
                _selectedConfiguration = value;
                OnPropertyChanged("SelectedConfiguration");
            }
        }
        private XElement _activeConfiguration;
        public XElement ActiveConfiguration
        {
            get { return _activeConfiguration; }
            set
            {
                _activeConfiguration = value;
                OnPropertyChanged("ActiveConfiguration");
                OnPropertyChanged("IsActiveConnectionSet");
            }
        }

        //Solutions
        public ObservableCollection<Solution> UnmanagedSolutions { get; set; }
        private Solution _selectedSolution;
        public Solution SelectedSolution 
        {
            get 
            { 
                return _selectedSolution; 
            }
            set
            {
                _selectedSolution = value;
                OnPropertyChanged("SelectedSolution");
            }
        }
        private Solution _activeSolution;
        public Solution ActiveSolution
        {
            get
            {
                return _activeSolution;
            }
            set
            {
                _activeSolution = value;
                OnPropertyChanged("ActiveSolution");
                OnPropertyChanged("IsActiveSolutionSet");
            }
        }

        //Active Publisher
        private Publisher _activePublisher;
        public Publisher ActivePublisher
        {
            get { return _activePublisher; }
            set { _activePublisher = value; OnPropertyChanged("ActivePublisher"); }
        }

        #endregion

        #region Fields
        //CRM Data Provider
        private ConsolelessServerConnection _serverConnect;        
        private static OrganizationServiceProxy _serviceProxy;
        private static OrganizationServiceContext _orgContext;

        BackgroundWorker worker;
        
        #endregion

        #region Constructor

        public MainWindowViewModel()
        {
            XDocument xmlPackagesDocument = XDocument.Load(PACKAGES_FILENAME);
            XmlPackageData = xmlPackagesDocument.Element("UtilityRoot");

            XDocument xmlConfigurationsDocument = XDocument.Load(CONFIG_FILENAME);
            XmlConfigData = xmlConfigurationsDocument.Element("Configurations");

            Configurations = new ObservableCollection<XElement>();
            Packages = new ObservableCollection<XElement>();
            UnmanagedSolutions = new ObservableCollection<Solution>();
            CurrentFiles = new ObservableCollection<FileInfo>();
            CurrentFilesSelected = new ObservableCollection<FileInfo>();
            WebResourceInfos = new ObservableCollection<XElement>();
            WebResourceInfosSelected = new ObservableCollection<XElement>();

            //Begin loading the XML data
            LoadXmlData();

            TabControlSelectedIndex = 0;
            _progressMessage = new StringBuilder();
            _shouldPublishAllAfterUpload = true;

            //Set up the background worker to handle upload web resources. Helps
            //prevent the UI from locking up and can have a Console-like output window
            //with real-time display.
            worker = new BackgroundWorker();            
            worker.WorkerReportsProgress = true;
            worker.DoWork += new DoWorkEventHandler(BeginUpload);
            worker.RunWorkerCompleted += delegate(object s, RunWorkerCompletedEventArgs args)
            {
                this.BeginInvoke(() =>
                {
                    AreAllButtonsEnabled = true;
                });
            };
        }

        #endregion

        #region Methods 

        private void LoadXmlData()
        {
            LoadConfigurations();
            LoadPackages();           
        }
     
        //Configuration Methods
        private void LoadConfigurations()
        {
            Configurations.Clear();

            var configs = XmlConfigData.Descendants("Configuration");
            foreach (var c in configs)
            {
                Configurations.Add(c);
            }
        }
        private void SaveConfigurations()
        {
            XmlConfigData.Descendants("Configuration").Remove();
            XmlConfigData.Add(Configurations.ToArray());

            XmlConfigData.Save(CONFIG_FILENAME);
        }
        private void CreateNewConfiguration()
        {
            XElement newConfig = new XElement("Configuration",
                new XAttribute("name", "New Connection"),
                new XAttribute("server", String.Empty),
                new XAttribute("orgName", String.Empty),
                new XAttribute("userName", String.Empty),
                new XAttribute("domain", String.Empty));

            Configurations.Add(newConfig);
            SelectedConfiguration = Configurations[Configurations.Count - 1];
        }
        private void DeleteSelectedConfiguration()
        {
            if (SelectedConfiguration != null)
            {
                //if trying to delete the configuration that is also active already,
                //let them by clearing ActiveConfiguration and solutions.
                if (SelectedConfiguration == ActiveConfiguration)
                {
                    ClearActiveConfiguration();
                    ClearSolutions();
                }

                //Finally clear the SelectedConfiguration and remove it from the list of Configurations.
                var toBeDeleted = Configurations.Where(x => x == SelectedConfiguration).FirstOrDefault();
                if (toBeDeleted != null)
                {
                    Configurations.Remove(toBeDeleted);
                    SelectedConfiguration = null;
                }
            }
        }
        private void ClearActiveConfiguration()
        {
            ActiveConfiguration = null;
        }
        private void ActivateSelectedConfiguration()
        {
            //User may have already been connected to another org, disconnect them.
            ClearActiveConfiguration();

            //Clear out any Solutions from the Solutions collection since they are 
            //Configuration specfic.
            ClearSolutions();

            //Instantiate new proxy. if it is successful, it will also
            //set the ActiveConfiguration and retrieve Solutions.
            InstantiateService();
        }
        private bool CanActivateSelectedConfiguration()
        {
            if (SelectedConfiguration != null &&
                !String.IsNullOrWhiteSpace(SelectedConfiguration.Attribute("server").Value) &&
                !String.IsNullOrWhiteSpace(SelectedConfiguration.Attribute("orgName").Value))
            {
                return true;
            }
            return false;
        }
    
        //Solution Methods
        private void LoadSolutions()
        {
            //Check whether it already exists
            QueryExpression queryUnmanagedSolutions = new QueryExpression
            {
                EntityName = Solution.EntityLogicalName,
                ColumnSet = new ColumnSet(true),
                Criteria = new FilterExpression()
            };
            queryUnmanagedSolutions.Criteria.AddCondition("ismanaged", ConditionOperator.Equal, false);
            EntityCollection querySolutionResults = _serviceProxy.RetrieveMultiple(queryUnmanagedSolutions);

            if (querySolutionResults.Entities.Count > 0)
            {
                //The Where() is important because a query for all solutions
                //where Type=Unmanaged returns 3 solutions. The CRM UI of a 
                //vanilla instance shows only 1 unmanaged solution: "Default". 
                //Assume "Active" and "Basic" should not be touched?
                UnmanagedSolutions = new ObservableCollection<Solution>(
                    querySolutionResults.Entities
                        .Select(x => x as Solution)
                        .Where(s => s.UniqueName != "Active" &&
                                    s.UniqueName != "Basic"
                        )
                    );

                //If only 1 solution returns just go ahead and default it.
                if (UnmanagedSolutions.Count == 1 && UnmanagedSolutions[0].UniqueName == "Default")
                {
                    SelectedSolution = UnmanagedSolutions[0];
                    ActiveSolution = SelectedSolution;

                    SetActivePublisher();

                    //Advance the user to the Packages TabItem
                    TabControlSelectedIndex = 2;
                }
                else
                {
                    //Advance the user to the Solutions TabItem
                    TabControlSelectedIndex = 1;
                }

                OnPropertyChanged("UnmanagedSolutions");
                OnPropertyChanged("SelectedSolution");
                OnPropertyChanged("ActiveSolution");
                
            }
        }
        private void InstantiateService()
        {
            try
            {
                if (SelectedConfiguration == null)
                    throw new Exception("Please choose a configuration.");

                //Get the Password
                string password = String.Empty;
                PasswordWindow pw = new PasswordWindow();
                pw.Owner = Application.Current.MainWindow;
                bool? submitted = pw.ShowDialog();
                if (submitted.Value)
                {
                    password = pw.GetPassword();
                }
                else
                {
                    ErrorWindow needPassword = new ErrorWindow("You need to supply a Password and Submit. Try again.");
                    needPassword.Owner = Application.Current.MainWindow;
                    needPassword.ShowDialog();
                    return;
                }
                

                _serverConnect = new ConsolelessServerConnection();
                ConsolelessServerConnection.Configuration _config = new ConsolelessServerConnection.Configuration();
                _config = _serverConnect.GetServerConfiguration(
                    SelectedConfiguration.Attribute("server").Value,
                    SelectedConfiguration.Attribute("orgName").Value,
                    SelectedConfiguration.Attribute("userName").Value, 
                    password,
                    SelectedConfiguration.Attribute("domain").Value);

                _serviceProxy = new OrganizationServiceProxy(_config.OrganizationUri,
                    _config.HomeRealmUri,
                    _config.Credentials,
                    _config.DeviceCredentials);

                // This statement is required to enable early-bound type support.
                _serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors
                    .Add(new ProxyTypesBehavior());

                // The OrganizationServiceContext is an object that wraps the service
                // proxy and allows creating/updating multiple records simultaneously.
                _orgContext = new OrganizationServiceContext(_serviceProxy);

                //Set the ActiveConnection
                ActiveConfiguration = SelectedConfiguration;

                //If all worked, retrieve the solutions.
                LoadSolutions();

            }
            catch (Exception e)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(e.Message);
                sb.AppendLine();
                sb.AppendLine("Please fix the Connection information and try again.");
                ErrorWindow errorWindow = new ErrorWindow(sb.ToString());
                errorWindow.Owner = Application.Current.MainWindow;
                var x = errorWindow.ShowDialog();
            }
        }
        private void ClearSolutions()
        {
            //Clear solutions
            UnmanagedSolutions.Clear();
            SelectedSolution = null;
            ActiveSolution = null;

            ActivePublisher = null;
        }   
        private void ActivateSelectedSolution()
        {
            if (SelectedSolution != null)
            {
                ActiveSolution = SelectedSolution;

                SetActivePublisher();

                OnPropertyChanged("ActiveSolution");

                //Advance the user to the Packages TabItem
                TabControlSelectedIndex = 2;
            }
        }
        private void SetActivePublisher()
        {
            if (ActiveSolution == null)
                return;

            var pub = from p in _orgContext.CreateQuery<Publisher>()
                        where p.PublisherId.Value == ActiveSolution.PublisherId.Id
                        select new Publisher
                        {
                            CustomizationPrefix = p.CustomizationPrefix
                             
                        };

            ActivePublisher = pub.First();
            OnPropertyChanged("ActivePublisher");

        }

        //Package Methods   
        private void LoadPackages()
        {
            Packages.Clear();
            var packages = XmlPackageData.Element("Packages").Descendants("Package");
            foreach (var p in packages)
            {
                Packages.Add(p);
            }
            OnPropertyChanged("Packages");
        }
        private void SavePackages()
        {
            //The user is influenced to believe a Save event will only 
            //save the ActivePackage but really it will save all of them.
            //Code is in place to prevent the user from editing one package then 
            //trying to load another without saving the first.

            //At this point the XmlRootData object is stale and needs to be 
            //repopulated with the Packages collection.
            XmlPackageData.Descendants("Package").Remove();

            //But the ActivePackage may have its Web Resources modified and they
            //need to be added back to the ActivePackage.
            if (ActivePackage != null)
            {
                ActivePackage.Elements("WebResourceInfo").Remove();
                ActivePackage.Add(WebResourceInfos.ToArray());
            }

            XmlPackageData.Element("Packages").Add(Packages.ToArray());

            XmlPackageData.Save(PACKAGES_FILENAME);

            IsActivePackageDirty = false;

        }
        private void DeleteSelectedPackage()
        {
            if (SelectedPackage != null)
            {
                var toBeDeleted = Packages.Where(x => x == SelectedPackage).FirstOrDefault();

                if (toBeDeleted != null)
                {
                    if (ActivePackage == SelectedPackage)
                    {
                        ActivePackage = null;
                        //Also, clear out any dependencies
                        CurrentFiles.Clear();
                        CurrentFilesSelected.Clear();
                        WebResourceInfos.Clear();
                        WebResourceInfosSelected.Clear();
                    }
                    Packages.Remove(toBeDeleted);
                    SelectedPackage = null;
                }
                SavePackages();
            }
        }
        private void ActivatePackage()
        {
            //Don't allow them to load a package without first saving
            //the ActivePackage if its dirty.
            if (ActivePackage != null && IsActivePackageDirty)
            {
                ErrorWindow dirtyPackageWindow =
                    new ErrorWindow("You have unsaved changes to the Active Package. Please save before loading another package.");
                dirtyPackageWindow.Owner = Application.Current.MainWindow;
                dirtyPackageWindow.ShowDialog();
                return;              
            }

            if (SelectedPackage != null)
            {
                ActivePackage = SelectedPackage;

                //Readies the Files DataGrid 
                SearchAndPopulateFiles();

                //Readies the Web Resources DataGrid
                LoadWebResourceInfos();
            }
        }
        private void CreateNewPackage()
        {
            if (ActivePackage != null)
            {
                SavePackages();
            }

            XElement newPackage = new XElement("Package",
                new XAttribute("name", "NewPackage"),
                new XAttribute("rootPath", String.Empty),
                new XAttribute("isNamePrefix", true));

            Packages.Add(newPackage);
            SelectedPackage = Packages[Packages.Count - 1];

            ActivatePackage();

            SavePackages();
        }
        private void LoadWebResourceInfos()
        {
            if (ActivePackage == null)
                return;

            //As always, clear the collection first.
            WebResourceInfos.Clear();
            WebResourceInfosSelected.Clear();

            var webResourceInfos = ActivePackage.Elements("WebResourceInfo");

            if (webResourceInfos != null)
            {
                foreach (var wr in webResourceInfos)
                {
                    WebResourceInfos.Add(wr);
                }
            }
        }
        private void SearchAndPopulateFiles()
        {
            if (ActivePackage == null)
                return;

            string searchText = FileSearchText; //Find all files

            string rootPath = ActivePackage.Attribute("rootPath").Value;
            DiscoverFiles(rootPath, searchText);

        }

        //Misc
        private void ShowBrowseFolderDialog()
        {
            System.Windows.Forms.FolderBrowserDialog dlgWRFolder = new System.Windows.Forms.FolderBrowserDialog()
            {
                Description = "Select the folder containing the potential Web Resource files",
                ShowNewFolderButton = false
            };
            if (dlgWRFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {                
                IsActivePackageDirty = false;

                //Because Web Resources are relative to the root path, 
                //all the current Web ResourceInfos should be cleared
                WebResourceInfos.Clear();
                WebResourceInfosSelected.Clear();

                //Change the rootpath and notify all bindings
                ActivePackage.Attribute("rootPath").Value = dlgWRFolder.SelectedPath;
                OnPropertyChanged("ActivePackage");

                //Auto-save
                SavePackages();

                //Display new files
                SearchAndPopulateFiles();
            }
        }
        private void DiscoverFiles(string rootPath, string searchText)
        {
            CurrentFiles.Clear();
            CurrentFilesSelected.Clear();

            if (rootPath != String.Empty)
            {
                DirectoryInfo di = new DirectoryInfo(rootPath);

                var files = di.EnumerateFiles("*", SearchOption.AllDirectories)
                    .Where(f => ResourceExtensions.ValidExtensions.Contains(f.Extension))
                    .Where(f => f.Name != PACKAGES_FILENAME);                

                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    files = files.Where(f => f.FullName.Contains(searchText));
                }

                foreach (FileInfo f in files)
                {
                    CurrentFiles.Add(f);                    
                }
                OnPropertyChanged("CurrentFiles");
            }
        }
        private void AddFilesToWebResources(object parameter)
        {
            //Set the ActivePackage as Dirty. 
            IsActivePackageDirty = true;

            //Clear the collection of selected files
            CurrentFilesSelected.Clear();

            //List<FileInfo> selectedFiles = new List<FileInfo>();
            if (parameter != null && parameter is IEnumerable)
            {
                foreach (var fileInfo in (IEnumerable)parameter)
                {
                    CurrentFilesSelected.Add((FileInfo)fileInfo);
                }
            }

            if (CurrentFilesSelected.Count > 0)
            {
                foreach (FileInfo fi in CurrentFilesSelected)
                {
                    //Add it to the list of web resource info, if not already there.
                    //The matching criteria will be the ?
                    
                    XElement newInfo = ConvertFileInfoToWebResourceInfo(fi);

                    if (WebResourceInfos.Where(w => w.Attribute("filePath").Value == newInfo.Attribute("filePath").Value).Count() == 0)
                    {
                        WebResourceInfos.Add(newInfo);                        
                    }
                    else
                    {
                        //it's already in the list! do nothing.
                    }
                }                
            }
        }
        private void DeleteSelectedWebResources(object parameter)
        {
            //Set the ActivePackage as Dirty. 
            IsActivePackageDirty = true;

            WebResourceInfosSelected.Clear();

            if (parameter != null && parameter is IEnumerable)
            {
                //Lists allow the ForEach extension method good for
                //removing items of a collection. Looping through an 
                //enumerable caused indexing errors after the first
                //iteration.
                List<XElement> infosToDelete = new List<XElement>();

                foreach (var wr in (IEnumerable)parameter)
                {
                    infosToDelete.Add((XElement)wr);
                }

                infosToDelete.ForEach(info => WebResourceInfos.Remove(info));
            }
        }
        private XElement ConvertFileInfoToWebResourceInfo(FileInfo fi)
        {
            var x = fi.Extension.Split('.');
            string type = x[x.Length - 1].ToLower();

            String name = fi.FullName.Replace(ActivePackage.Attribute("rootPath").Value.Replace("/", "\\"), String.Empty);

            XElement newWebResourceInfo = new XElement("WebResourceInfo",
                new XAttribute("name", name.Replace("\\", "/")),
                new XAttribute("filePath", name),
                new XAttribute("displayName", fi.Name),
                new XAttribute("type", type),
                new XAttribute("description", String.Empty));

            return newWebResourceInfo;
        }

        private void BeginUpload(object s, DoWorkEventArgs args)
        {
            //Retrieve all Web Resources so we can determine if each
            //needs be created or updated.
            var crmResources = RetrieveWebResourcesForActiveSolution();

            //Create or Update the WebResource
            if (WebResourceInfosSelected.Count > 0)
            {
                _progressMessage.Clear();
                AddMessage(String.Format("Processing {0} Web Resources...", WebResourceInfosSelected.Count.ToString()));
                int i = 1;

                foreach (XElement fi in WebResourceInfosSelected)
                {
                    string name = GetWebResourceFullNameIncludingPrefix(fi);
                    
                    AddMessage(String.Empty);
                    AddMessage(i.ToString() + ") " + name);

                    if (IsWebResourceNameValid(name))
                    {
                        //If the Unmanaged Solution already contains the Web Resource,
                        //do an Update.
                        var resourceThatMayExist = crmResources.Where(w => w.Name == name).FirstOrDefault();
                        if (resourceThatMayExist != null)
                        {
                            AddMessage("Already exists. Updating...");
                            UpdateWebResource(fi, resourceThatMayExist);
                            AddMessage("Done.");
                        }
                        //If not, create the Web Resource and a Solution Component.
                        else
                        {
                            AddMessage("Doesn't exist. Creating...");
                            CreateWebResource(fi);
                            AddMessage("Done.");
                        }
                    }
                    else
                    {
                        AddMessage(VALID_NAME_MSG);
                    }

                    i++;
                }
                
                AddMessage(String.Empty);
                AddMessage("Done processing files.");

                //All WebResources should be in. Publish all.
                if (ShouldPublishAllAfterUpload)
                {
                    AddMessage(String.Empty);
                    AddMessage("You chose to publish all customizations. Please be patient as it may take a few minutes to complete.");
                    PublishAll();
                }
                else
                {
                    AddMessage(String.Empty);
                    AddMessage("You chose not to publish all customizations.");
                }
                AddMessage("Process complete!");                
            }
        }
        private void UploadWebResources(object parameter)
        {
            IsOutputWindowDisplayed = true;
            AreAllButtonsEnabled = false;

            //Clear the collection of selected Web Resources
            WebResourceInfosSelected.Clear();

            if (parameter != null && parameter is IEnumerable)
            {
                foreach (var webResource in (IEnumerable)parameter)
                {
                    WebResourceInfosSelected.Add((XElement)webResource);
                }
            }
            worker.RunWorkerAsync(WebResourceInfosSelected);
        }        
        private void UploadAllWebResources()
        {
            UploadWebResources(WebResourceInfos);            
        }

        private bool CanUploadWebResource()
        {
            if (ActiveConfiguration != null &&
                ActiveSolution != null &&
                ActivePublisher != null &&
                ActivePackage != null)
            {
                return true;
            }
            return false;
        }
        private void PublishAll()
        {
            try
            {
                AddMessage(String.Empty);
                AddMessage("Publishing all customizations...");     
                
                PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
                var response = (PublishAllXmlResponse)_serviceProxy.Execute(publishRequest);

                AddMessage("Done.");
            }
            catch (Exception e)
            {
                AddMessage("Error publishing: " + e.Message);
            }
        }        
        private IEnumerable<WebResource> RetrieveWebResourcesForActiveSolution()
        {
            //The following query finds all WebResources that are SolutionComponents for
            //the ActiveSolution. Simply querying WebResources does not retrieve the desired
            //results. Additionally, when creating WebResources, you must create a SolutionComponent
            //if attaching to any unmanaged solution other than the "Default Solution."
            var webResources = from wr in _orgContext.CreateQuery<WebResource>()
                               join sc in _orgContext.CreateQuery<SolutionComponent>()
                                 on wr.WebResourceId equals sc.ObjectId
                               where wr.IsManaged == false
                               where wr.IsCustomizable.Value == true
                               where sc.ComponentType.Value == (int)componenttype.WebResource
                               where sc.SolutionId.Id == ActiveSolution.SolutionId.Value
                               select new WebResource
                               {
                                     WebResourceType = wr.WebResourceType,
                                     WebResourceId = wr.WebResourceId,
                                     DisplayName = wr.DisplayName,
                                     Name = wr.Name,
                                     // Content = wr.Content, Removed to improve performance
                                     Description = wr.Description
                               };

            return webResources.AsEnumerable();
        }
        private void CreateWebResource(XElement webResourceInfo)
        {
            
                try
                {
                    //Create the Web Resource.
                    WebResource wr = new WebResource()
                    {
                        Content = getEncodedFileContents(ActivePackage.Attribute("rootPath").Value + webResourceInfo.Attribute("filePath").Value),
                        DisplayName = webResourceInfo.Attribute("displayName").Value,
                        Description = webResourceInfo.Attribute("description").Value,
                        LogicalName = WebResource.EntityLogicalName,
                        Name = GetWebResourceFullNameIncludingPrefix(webResourceInfo)

                    };

                    wr.WebResourceType = new OptionSetValue((int)ResourceExtensions.ConvertStringExtension(webResourceInfo.Attribute("type").Value));

                    //Special cases attributes for different web resource types.
                    switch (wr.WebResourceType.Value)
                    {
                        case (int)ResourceExtensions.WebResourceType.Silverlight:
                            wr.SilverlightVersion = "4.0";
                            break;
                    }

                    // ActivePublisher.CustomizationPrefix + "_/" + ActivePackage.Attribute("name").Value + webResourceInfo.Attribute("name").Value.Replace("\\", "/"),
                    Guid theGuid = _serviceProxy.Create(wr);

                    //If not the "Default Solution", create a SolutionComponent to assure it gets
                    //associated with the ActiveSolution. Web Resources are automatically added
                    //as SolutionComponents to the Default Solution.
                    if (ActiveSolution.UniqueName != "Default")
                    {
                        AddSolutionComponentRequest scRequest = new AddSolutionComponentRequest();
                        scRequest.ComponentType = (int)componenttype.WebResource;
                        scRequest.SolutionUniqueName = ActiveSolution.UniqueName;
                        scRequest.ComponentId = theGuid;
                        var response = (AddSolutionComponentResponse)_serviceProxy.Execute(scRequest);
                    }
                }
                catch (Exception e)
                {
                    AddMessage("Error: " + e.Message);
                    return;
                }            
            
        }
        private void UpdateWebResource(XElement webResourceInfo, WebResource existingResource)
        {
            try
            {
                //These are the only 3 things that should (can) change.
                WebResource wr = new WebResource()
                {
                    Content = getEncodedFileContents(ActivePackage.Attribute("rootPath").Value + webResourceInfo.Attribute("filePath").Value),
                    DisplayName = webResourceInfo.Attribute("displayName").Value,
                    Description = webResourceInfo.Attribute("description").Value
                };
                wr.WebResourceId = existingResource.WebResourceId;
                _serviceProxy.Update(wr);

            }
            catch (Exception e)
            {
                AddMessage("Error: " + e.Message);
                return;
            }
        }
        private string getEncodedFileContents(String pathToFile)
        {
            FileStream fs = new FileStream(pathToFile, FileMode.Open, FileAccess.Read);
            byte[] binaryData = new byte[fs.Length];
            long bytesRead = fs.Read(binaryData, 0, (int)fs.Length);
            fs.Close();
            return System.Convert.ToBase64String(binaryData, 0, binaryData.Length);
        }
        private string GetWebResourceFullNameIncludingPrefix(XElement webResourceInfo)
        {
            //The Web Resource name always starts with the Publisher's Prefix
            //i.e., "new_"
            string name = ActivePublisher.CustomizationPrefix + "_";

            //Check to see if the user has chosen to add the Package Name as part of the 
            //prefix.
            if (!String.IsNullOrWhiteSpace(ActivePackage.Attribute("isNamePrefix").Value) &&
                Boolean.Parse(ActivePackage.Attribute("isNamePrefix").Value) == true)
            {
                name += "/" + ActivePackage.Attribute("name").Value;
            }
                 
            //Finally add the name on to the prefix
            name += webResourceInfo.Attribute("name").Value;

            return name;
        }
        private bool IsWebResourceNameValid(string name)
        {
            Regex inValidWRNameRegex = new Regex("[^a-z0-9A-Z_\\./]|[/]{2,}", 
                (RegexOptions.Compiled | RegexOptions.CultureInvariant));

            bool result = true;

            //Test valid characters
            if (inValidWRNameRegex.IsMatch(name))
            {
                AddMessage(VALID_NAME_MSG);
                result = false;
            }

            //Test length
            //Remove the customization prefix and leading _ 
            if (name.Remove(0, ActivePublisher.CustomizationPrefix.Length + 1).Length > 100)
            {
                AddMessage("ERROR: Web Resource name must be <= 100 characters.");
                result = false;
            }
  
            return result;
        }
        private void AddMessage(string msg)
        {
            //Assures that this happens on the UI thread
            this.BeginInvoke(() =>
            {
                ProgressMessage = msg;
            });
        }

        #endregion

        
    }
}
