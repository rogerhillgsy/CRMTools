using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Query;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Linq;
using System.Collections.Generic;
using System.Configuration;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.KeyVault.Models;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Threading.Tasks;
using Microsoft.Xrm.Tooling.Connector;

namespace ImportContact
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string errorFilePath = string.Empty;
        Boolean isErrorRecord = false;
        int numberOfDuplicateRecords = 0;
        int numberOfValidationError = 0;
        int numberOfRecordsImported = 0;
        int numberOfRecordsUpdated = 0;
        OptionSetMetadata optionSetMetatdataBusinessInterest = new OptionSetMetadata();


        const string CLIENTSECRET = "Z~-V7ccMO-y7-J_Y.9ddKvM7q0Ei2B04SZ";
        const string CLIENTID = "6a79f939-19d4-45fa-8db2-ed36a6884acf";
        string KeyVaultURI = ConfigurationManager.AppSettings["KeyVaultPath"] + "/secrets/CrmHub-Password/fcf20dcc597a40de845342beff48bff6";

        static KeyVaultClient kvc = null;  
        static string password = string.Empty;

        private void Button1_Click(object sender, EventArgs e)
        {

            KeyVaultClient kvc = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(GetToken));
           // var secret = kvc.GetSecretAsync("https://crmsecrets.vault.azure.net/secrets/CrmHub-Password/fcf20dcc597a40de845342beff48bff6");
            var secret = kvc.GetSecretAsync(KeyVaultURI);
            password = secret.Result.Value;
         

            //string filePath = string.Empty;
            string fileExt = string.Empty;
            OpenFileDialog openFileDialogExcel = new OpenFileDialog(); //open dialog to choose file  
            openFileDialogExcel.Title = "Browse Excel Files";

            openFileDialogExcel.CheckFileExists = true;
            openFileDialogExcel.CheckPathExists = true;

            openFileDialogExcel.DefaultExt = "txt";
            openFileDialogExcel.Filter = "xlsx Files (*.xlsx)|*.xlsx |All files (*.*)|*.*";
            openFileDialogExcel.FilterIndex = 2;
            openFileDialogExcel.RestoreDirectory = true;
            if (openFileDialogExcel.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
            {
                txtFilePath.Text = openFileDialogExcel.FileName; //get the path of the file  

            }
            else
            {
                MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
            }
        }


        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=Yes';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select DISTINCT * from [Data Sheet$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  

           
                }
                catch (Exception ex)
                {
                    throw (ex);
                }
            }
            return RemoveEmptyRows(dtexcel);
        }

        private DataTable RemoveEmptyRows(DataTable dt)
        {
            List<int> rowIndexesToBeDeleted = new List<int>();
            int indexCount = 0;
            foreach (var row in dt.Rows)
            {
                var r = (DataRow)row;
                int emptyCount = 0;
                int itemArrayCount = r.ItemArray.Length;
                foreach (var i in r.ItemArray) if (string.IsNullOrWhiteSpace(i.ToString())) emptyCount++;

                if (emptyCount == itemArrayCount) rowIndexesToBeDeleted.Add(indexCount);

                indexCount++;
            }

            int count = 0;
            foreach (var i in rowIndexesToBeDeleted)
            {
                dt.Rows.RemoveAt(i - count);
                count++;
            }

            return dt;
        }
        public Boolean ValidateFieldNullOrEmpty(string excelColumnName, DataTable dtExcel, DataRow drExcel, DataTable dtError, Boolean isUpdateRequest = false)
        {
            Boolean validationSuccess;
            if (dtExcel.Columns.Contains(excelColumnName) && !String.IsNullOrEmpty(drExcel[excelColumnName].ToString()))
            {             
                validationSuccess = true;
            }
            else
            {
                if (!isUpdateRequest)
                {
                    PopulateErrorTable(drExcel["Email Address"].ToString(), excelColumnName, string.Format("{0} is required Field to create contact", excelColumnName), dtError);
                }
                validationSuccess = false;
            }
            return validationSuccess;
        }
        public void RequiredFieldValidationAndAssignment(IOrganizationService service, DataTable dtExcel, DataRow drExcel, Entity contact, DataTable dtError, Boolean isUpdateRequest = false)
        {
            Guid contactId = GetContactByEmail(service, drExcel["Email Address"].ToString());
            if ((contactId == Guid.Empty && !isUpdateRequest) || (contactId != Guid.Empty && isUpdateRequest))
            {
                if (isUpdateRequest)
                {
                    contact["contactid"] = contactId; // to update contact , uniqueid will be required. No nee to add emailaddress in cntact collection to update it.
                }
                else
                {
                    contact["emailaddress1"] = drExcel["Email Address"];
                }

               if (ValidateFieldNullOrEmpty("First Name", dtExcel, drExcel, dtError, isUpdateRequest))
                {
                    contact["firstname"] = drExcel["First Name"];
            }
            if (ValidateFieldNullOrEmpty("Last Name", dtExcel, drExcel, dtError, isUpdateRequest))
            {
                contact["lastname"] = drExcel["Last Name"];
            }

            if (ValidateFieldNullOrEmpty("Current Organisation", dtExcel, drExcel, dtError, isUpdateRequest))
            {
                contact["arup_currentorganisation"] = drExcel["Current Organisation"];
            }
            if (ValidateFieldNullOrEmpty("Town/City", dtExcel, drExcel, dtError, isUpdateRequest))
            {
                contact["address1_city"] = drExcel["Town/City"];
            }

            if (ValidateFieldNullOrEmpty("Role", dtExcel, drExcel, dtError, isUpdateRequest))
            {
                AssignRole(drExcel, contact, dtError);
            }
            if (ValidateFieldNullOrEmpty("Country", dtExcel, drExcel, dtError, isUpdateRequest))
            {
                AssignCountry(dtExcel, drExcel, contact, dtError, service);
            }

            if (ValidateFieldNullOrEmpty("Business Interest", dtExcel, drExcel, dtError, isUpdateRequest))
            {
                AssignBusinessInterest(drExcel, contact, service);
                }
            }
            else
            {
                if (!isUpdateRequest)
                {
                    PopulateErrorTable(drExcel["Email Address"].ToString(), "Email Addresss", string.Format("This is duplicate record.Contact with Email Address {0} already exist", drExcel["Email Address"].ToString()), dtError);
                    numberOfDuplicateRecords++;
                }
                else
                {
                    PopulateErrorTable(drExcel["Email Address"].ToString(), "Email Addresss", string.Format("This Email Address {0} doesn't exist in CRM to update record", drExcel["Email Address"].ToString()), dtError);
                }
            }


        }

        public string GetCRMConnectionString()
        {
         
            string connectionString = "";
            

            CrmConnection crmConnection = new CrmConnection();
            if (cmbEnvironment.SelectedItem.ToString().ToUpper() == "UAT".ToUpper())
                connectionString = ConfigurationManager.ConnectionStrings["CRMUAT"].ConnectionString;
            else if (cmbEnvironment.SelectedItem.ToString().ToUpper() == "PRODUCTION".ToUpper())
                connectionString = ConfigurationManager.ConnectionStrings["CRMPROD"].ConnectionString;

            connectionString = connectionString.Replace("%Password%", password);

            return connectionString;
        }



        public static async Task<string> GetToken(string authority, string resource, string scope)
        {
            var authContext = new AuthenticationContext(authority);
            ClientCredential clientCred = new ClientCredential(CLIENTID, CLIENTSECRET);
            AuthenticationResult result = await authContext.AcquireTokenAsync(resource, clientCred);

            if (result == null)
                throw new InvalidOperationException("Failed to obtain the JWT token");

            return result.AccessToken;
        }
        public void CreateUpdateContact(DataTable dtExcel, DataTable dtError)
        { 
        // CrmConnection con = new CrmConnection("CRM");
        // IOrganizationService service = new OrganizationService(GetCRMConnectionString());
             var CrmService = new CrmServiceClient(GetCRMConnectionString());
            IOrganizationService service = CrmService.OrganizationServiceProxy;

            progressBar1.Visible = true;
            progressBar1.Minimum = 1;
            progressBar1.Maximum = dtExcel.Rows.Count;
            progressBar1.Value = 1;
            progressBar1.Step = 1;
            Boolean isCreateRequest = rdBtnCreateContact.Checked;
            Boolean isUpdateRequest = rdBtnUpdateContact.Checked;

            foreach (DataRow drExcel in dtExcel.Rows)
            {
                isErrorRecord = false;
                Entity contact = new Entity();
                if (isCreateRequest)
                {
                    if (ValidateFieldNullOrEmpty("Email Address", dtExcel, drExcel, dtError,false))
                    {
                        RequiredFieldValidationAndAssignment(service, dtExcel, drExcel, contact, dtError, false);
                        if (!isErrorRecord)
                        {
                            OptionalFieldValidationAndAssignment(service, dtExcel, drExcel, contact,dtError);
                            contact["arup_contacttype"] = new OptionSetValue(FieldConstant.OPTIONSET_CONTACTTYPE_VALUE_MARKETING);
                            contact.LogicalName = "contact";
                            service.Create(contact);
                            numberOfRecordsImported++;
                        }
                    }
                }
                else if (isUpdateRequest)
                {
                    if (ValidateFieldNullOrEmpty("Email Address", dtExcel, drExcel, dtError,true))
                    {
                        RequiredFieldValidationAndAssignment(service, dtExcel, drExcel, contact, dtError,true);
                        if (contact.Attributes.Contains("contactid"))
                        {
                            OptionalFieldValidationAndAssignment(service, dtExcel, drExcel, contact,dtError);
                            contact["arup_contacttype"] = new OptionSetValue(FieldConstant.OPTIONSET_CONTACTTYPE_VALUE_MARKETING);
                            contact.LogicalName = "contact";
                            service.Update(contact);
                            numberOfRecordsUpdated++;
                        }
                    }

                }
                lblProgressBar.Text = "Processing Records......";
                progressBar1.PerformStep();
            }
        }

        public void AssignRole(DataRow drExcel, Entity contact, DataTable dtError)
        {
            int roleOptionSetValue = GetOptionSetValue(drExcel["Role"].ToString());
            if (roleOptionSetValue != 0)
            {
                contact["accountrolecode"] = new OptionSetValue(roleOptionSetValue);
            }
            else
            {
                PopulateErrorTable(drExcel["Email Address"].ToString(), "Role", string.Format("Role {0} doesn't exist. Valid value for roles are {1},{2},{3},{4}", drExcel["Role"].ToString(), FieldConstant.OPTIONSET_ROLE_KEY_DECISIONMAKER, FieldConstant.OPTIONSET_ROLE_KEY_INFLUENCER, FieldConstant.OPTIONSET_ROLE_KEY_EMPLOYEE, FieldConstant.OPTIONSET_ROLE_VALUE_CSUITE), dtError);
            }
        }
        public void AssignCountry(DataTable dtExcel, DataRow drExcel, Entity contact, DataTable dtError, IOrganizationService service)
        {
            string countryName = drExcel["Country"].ToString();
            Guid countryId = GetCountryByName(service, countryName);
            if (countryId != Guid.Empty)
            {

                contact["ccrm_countryid"] = new EntityReference("ccrm_country", countryId);

                if (countryName.ToUpper() == FieldConstant.COUNTRY_AUSTRALIA || countryName.ToUpper() == FieldConstant.COUNTRY_CANADA || countryName.ToUpper() == FieldConstant.COUNTRY_INDONESIA ||
                    countryName.ToUpper() == FieldConstant.COUNTRY_MALAYSIA || countryName.ToUpper() == FieldConstant.COUNTRY_NEW_ZEALAND || countryName.ToUpper() == FieldConstant.COUNTRY_SINGAPORE ||
                    countryName.ToUpper() == FieldConstant.COUNTRY_USA)
                {
                    if (ValidateFieldNullOrEmpty("State/Province", dtExcel, drExcel, dtError))
                    {
                        Guid stateId = GetStateByName(service, drExcel["State/Province"].ToString());
                        if (stateId != Guid.Empty)
                        {

                            contact["ccrm_countrystate"] = new EntityReference("ccrm_arupusstate", stateId);
                        }
                        else
                        {
                            PopulateErrorTable(drExcel["Email Address"].ToString(), "State/Province", string.Format("State/Province - {0} doesn't exist. Please provide valid State/Province", drExcel["State/Province"].ToString()), dtError);
                        }
                    }
                }else
                {
                    //optional field without validation
                    if (dtExcel.Columns.Contains("State/Province") && !string.IsNullOrEmpty(drExcel["State/Province"].ToString()))
                        contact["address1_stateorprovince"] = drExcel["State/Province"];
                }

            }
            else
            {
                PopulateErrorTable(drExcel["Email Address"].ToString(), "Country", string.Format("Country {0} doesn't exist. Please provide valid country name", drExcel["Organisation name"].ToString()), dtError);
            }
        }
        public void AssignBusinessInterest(DataRow drExcel, Entity contact, IOrganizationService service)
        {
            string[] arrBusniessInterest = drExcel["Business Interest"].ToString().Split(',');
            if (optionSetMetatdataBusinessInterest.Options.Count <= 0)
            {
                optionSetMetatdataBusinessInterest = GetOptionSetMetaData(service, "contact", "ccrm_businessinterest");
            }

            int optionSetValueBusinessInterest = 0;
            string picklistNameBusinessInterest = string.Empty;
            string picklistValueBusinessInterest = string.Empty;

            int counter = 0;
            foreach (string item in arrBusniessInterest)
            {
                counter++;
                var optionSetValue = (from o in optionSetMetatdataBusinessInterest.Options
                                      where o.Label.UserLocalizedLabel.Label.ToLower() == item.ToLower()
                                      select o.Value).FirstOrDefault();

                if (counter == 1)
                {
                    optionSetValueBusinessInterest = Convert.ToInt32(optionSetValue);
                    picklistNameBusinessInterest = item;
                    picklistValueBusinessInterest = optionSetValue.ToString();
                }
                else
                {
                    picklistNameBusinessInterest = string.Concat(picklistNameBusinessInterest, ",", item);
                    picklistValueBusinessInterest = string.Concat(picklistValueBusinessInterest, ",", optionSetValue.ToString());
                }
            }

            if (optionSetValueBusinessInterest > 0)
            {
                contact["ccrm_businessinterest"] = new OptionSetValue(Convert.ToInt32(optionSetValueBusinessInterest));
            }
            contact["ccrm_businessinterestpicklistname"] = picklistNameBusinessInterest;
            contact["ccrm_businessinterestpicklistvalue"] = picklistValueBusinessInterest;
        }
        public void OptionalFieldValidationAndAssignment(IOrganizationService service, DataTable dtExcel, DataRow drExcel, Entity contact,DataTable dtError)
        {
            if (dtExcel.Columns.Contains("Title") && !string.IsNullOrEmpty(drExcel["Title"].ToString()))
                contact["salutation"] = drExcel["Title"];

            if (dtExcel.Columns.Contains("Job Title") && !string.IsNullOrEmpty(drExcel["Job Title"].ToString()))
                contact["jobtitle"] = drExcel["Job Title"];

            if (dtExcel.Columns.Contains("Role") && !string.IsNullOrEmpty(drExcel["Role"].ToString()))
            {
                AssignRole(drExcel, contact, dtError);
            }

            //if (dtExcel.Columns.Contains("State/Province (free text)") && !string.IsNullOrEmpty(drExcel["State/Province (free text)"].ToString()))
            //    contact["address1_stateorprovince"] = drExcel["State/Province (free text)"];

            if (dtExcel.Columns.Contains("Entering on behalf of") && !string.IsNullOrEmpty(drExcel["Entering on behalf of"].ToString()))
            {
                Guid enteringOnBehalfOfId = GetUserByName(service, drExcel["Entering on behalf of"].ToString());
                if (enteringOnBehalfOfId != Guid.Empty)
                {
                    contact["preferredsystemuserid"] = new EntityReference("systemuser", enteringOnBehalfOfId);
                }
            }

            if (dtExcel.Columns.Contains("Owner") && !string.IsNullOrEmpty(drExcel["Owner"].ToString()))
            {
                Guid ownerId = GetUserByName(service, drExcel["Owner"].ToString());
                if (ownerId != Guid.Empty)
                {
                    contact["ownerid"] = new EntityReference("systemuser", ownerId);
                  //  contact["preferredsystemuserid"] = new EntityReference("systemuser", ownerId);
                }
            }

            if (dtExcel.Columns.Contains("Telephone") && !string.IsNullOrEmpty(drExcel["Telephone"].ToString()))
                contact["telephone1"] = drExcel["Telephone"];

            if (dtExcel.Columns.Contains("Mobile") && !string.IsNullOrEmpty(drExcel["Mobile"].ToString()))
                contact["mobilephone"] = drExcel["Mobile"].ToString();

            if (dtExcel.Columns.Contains("Department") && !string.IsNullOrEmpty(drExcel["Department"].ToString()))
                contact["department"] = drExcel["Department"];

            if (dtExcel.Columns.Contains("Local Given Name") && !string.IsNullOrEmpty(drExcel["Local Given Name"].ToString()))
                contact["ccrm_localgivenname"] = drExcel["Local Given Name"];

            if (dtExcel.Columns.Contains("Local Family Name") && !string.IsNullOrEmpty(drExcel["Local Family Name"].ToString()))
                contact["ccrm_localfamilyname"] = drExcel["Local Family Name"];

            if (dtExcel.Columns.Contains("Sync To Marketo") && !string.IsNullOrEmpty(drExcel["Sync To Marketo"].ToString()))
                contact["arup_synctomkto"] = drExcel["Sync To Marketo"].ToString().ToUpper() == "YES".ToString().ToUpper() ? true : false;

        }
        public OptionSetMetadata GetOptionSetMetaData(IOrganizationService service, string entityLogicalName, string optionSetAttributeName)
        {
            var attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityLogicalName,
                LogicalName = optionSetAttributeName,
                RetrieveAsIfPublished = true
            };

            var attributeResponse = (RetrieveAttributeResponse)service.Execute(attributeRequest);
            var attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

            OptionSetMetadata optionSetMetadata = attributeMetadata.OptionSet;
            //var optionList = (from o in attributeMetadata.OptionSet.Options
            //                  select new { Value = o.Value, Text = o.Label.UserLocalizedLabel.Label }).ToList();

            return optionSetMetadata;
        }
        public Guid GetCurrentOrganisationByName(IOrganizationService service, string currentOrganisation)
        {
            Guid organisationId = Guid.Empty;
            QueryExpression queryOrganisation = new QueryExpression("account");
            queryOrganisation.ColumnSet = new ColumnSet("accountid");
            queryOrganisation.Criteria.AddCondition(new ConditionExpression("name", ConditionOperator.Equal, currentOrganisation));

            EntityCollection collOrganisation = new EntityCollection();
            collOrganisation = service.RetrieveMultiple(queryOrganisation);

            if (collOrganisation.Entities.Count > 0)
            {
                organisationId = collOrganisation[0].Id;
            }

            return organisationId;
        }

        public Guid GetCountryByName(IOrganizationService service, string country)
        {
            Guid countryId = Guid.Empty;
            QueryExpression queryCountry = new QueryExpression("ccrm_country");
            queryCountry.ColumnSet = new ColumnSet("ccrm_countryid");
            queryCountry.Criteria.AddCondition(new ConditionExpression("ccrm_name", ConditionOperator.Equal, country));

            EntityCollection collCountry = new EntityCollection();
            collCountry = service.RetrieveMultiple(queryCountry);

            if (collCountry.Entities.Count > 0)
            {
                countryId = collCountry[0].Id;
            }

            return countryId;
        }

        public Guid GetStateByName(IOrganizationService service, string state)
        {
            Guid stateId = Guid.Empty;
            QueryExpression queryState = new QueryExpression("ccrm_arupusstate");
            queryState.ColumnSet = new ColumnSet("ccrm_arupusstateid");
            queryState.Criteria.AddCondition(new ConditionExpression("ccrm_name", ConditionOperator.Equal, state));

            EntityCollection collState = new EntityCollection();
            collState = service.RetrieveMultiple(queryState);

            if (collState.Entities.Count > 0)
            {
                stateId = collState[0].Id;
            }

            return stateId;
        }

        public Guid GetUserByName(IOrganizationService service, string userName)
        {
            Guid userId = Guid.Empty;
            QueryExpression queryUser = new QueryExpression("systemuser");
            queryUser.ColumnSet = new ColumnSet("systemuserid");
            queryUser.Criteria.AddCondition(new ConditionExpression("fullname", ConditionOperator.Equal, userName));

            EntityCollection collUser = new EntityCollection();
            collUser = service.RetrieveMultiple(queryUser);

            if (collUser.Entities.Count > 0)
                userId = collUser.Entities[0].Id;

            return userId;
        }
        public Guid GetContactByEmail(IOrganizationService service, string emailId)
        {
            Guid contactId = Guid.Empty;
            QueryExpression queryContact = new QueryExpression("contact");
            queryContact.ColumnSet = new ColumnSet("contactid");
            queryContact.Criteria.AddCondition(new ConditionExpression("emailaddress1", ConditionOperator.Equal, emailId));
            queryContact.Criteria.AddCondition(new ConditionExpression("statecode", ConditionOperator.Equal, 0));

            EntityCollection collContact = new EntityCollection();
            collContact = service.RetrieveMultiple(queryContact);

            if (collContact.Entities.Count > 0)
                contactId = collContact.Entities[0].Id;

            return contactId;
        }

        public int GetOptionSetValue(string optionSetKey)
        {
            int optionSetValue = 0;

            if (optionSetKey.Equals(FieldConstant.OPTIONSET_ROLE_KEY_DECISIONMAKER))
                optionSetValue = FieldConstant.OPTIONSET_ROLE_VALUE_DECISIONMAKER;
            else if (optionSetKey.Equals(FieldConstant.OPTIONSET_ROLE_KEY_EMPLOYEE))
                optionSetValue = FieldConstant.OPTIONSET_ROLE_VALUE_EMPLOYEE;
            else if (optionSetKey.Equals(FieldConstant.OPTIONSET_ROLE_KEY_INFLUENCER))
                optionSetValue = FieldConstant.OPTIONSET_ROLE_VALUE_INFLUENCER;
            else if (optionSetKey.Equals(FieldConstant.OPTIONSET_ROLE_KEY_CSUITE))
                optionSetValue = FieldConstant.OPTIONSET_ROLE_VALUE_CSUITE;

            return optionSetValue;
        }

        public void PopulateErrorTable(string columnFullName, string errorColumnName, string columnErrorDescription, DataTable dtError)
        {
            DataRow drError = dtError.NewRow();
            drError["Contact Name"] = columnFullName;
            drError["Error Column Name"] = errorColumnName;
            drError["Error Description"] = columnErrorDescription;
            dtError.Rows.Add(drError);

            isErrorRecord = true;
            numberOfValidationError++;
        }

        public void DisplaySummary(int numOfRecordProcessed, DataTable dtError)
        {
            grpBoxSummary.Visible = true;
            lblNumOfRecordProcessed.Visible = true;
            lblNumOfDuplicateRecords.Visible = rdBtnCreateContact.Checked ? true :false;
            lblTotalValidationError.Visible = true;
            lblTotalRecordsImported.Visible = true;

            lblNumOfRecordProcessed.Text = lblNumOfRecordProcessed.Text + ": " + numOfRecordProcessed;          
            lblTotalValidationError.Text = lblTotalValidationError.Text + ": " + numberOfValidationError;
            if (rdBtnCreateContact.Checked)
            {
                lblTotalRecordsImported.Text = lblTotalRecordsImported.Text + ": " + numberOfRecordsImported;
                lblNumOfDuplicateRecords.Text = lblNumOfDuplicateRecords.Text + ": " + numberOfDuplicateRecords;
            } else if(rdBtnUpdateContact.Checked){
                lblTotalRecordsImported.Text = "Total # Records Updated" + ": " + numberOfRecordsUpdated;
            }


        }
        public void ExportDataToExcel(DataTable dtErrorFile)
        {

            Excel.Application XlObj = new Excel.Application();
            XlObj.Visible = false;
            Excel._Workbook WbObj = (Excel.Workbook)(XlObj.Workbooks.Add(""));
            Excel._Worksheet WsObj = (Excel.Worksheet)WbObj.ActiveSheet;
            object misValue = System.Reflection.Missing.Value;


            try
            {
                int row = 1; int col = 1;
                foreach (DataColumn column in dtErrorFile.Columns)
                {
                    //adding columns
                    WsObj.Cells[row, col] = column.ColumnName;
                    col++;
                }
                //reset column and row variables
                col = 1;
                row++;
                for (int i = 0; i < dtErrorFile.Rows.Count; i++)
                {
                    //adding data
                    foreach (var cell in dtErrorFile.Rows[i].ItemArray)
                    {
                        WsObj.Cells[row, col] = cell;
                        col++;
                    }
                    col = 1;
                    row++;
                }
               string currentDate = DateTime.Now.ToString("MMddyyyy");
               errorFilePath = string.Format(@"C:\ErrorFolder\{0}.xlsx", "ErrorFile_" + currentDate);
            

                WbObj.SaveAs(errorFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                MessageBox.Show(string.Format("Log File is generated at {0}", errorFilePath));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                WbObj.Close(true, misValue, misValue);
            }
        }
      
        public Boolean FormInputValidation()
        {
            Boolean isValidationSuccess = true;
            string validationSummary = string.Empty;
            if (cmbEnvironment.SelectedIndex < 0)
            {
                isValidationSuccess = false;
                validationSummary =  "Please select the Environment in which records to be imported";
            }
            if (!rdBtnCreateContact.Checked && ! rdBtnUpdateContact.Checked)
            {
                isValidationSuccess = false;
                validationSummary = validationSummary + "\r\n" + "Please select Request Type";                     
            }
            if (string.IsNullOrEmpty(txtFilePath.Text))
            {
                isValidationSuccess = false;
                validationSummary = validationSummary + "\r\n" + "Please browse the excel file to be imported";            
            }
            if (!isValidationSuccess)
            {
                MessageBox.Show(validationSummary);
            }
            return isValidationSuccess;
        }
        private void btnOk_Click(object sender, EventArgs e)
        {
            if (FormInputValidation())
            {
                string fileExt = Path.GetExtension(txtFilePath.Text); //get the file extension  
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0 || fileExt.CompareTo(".xlsm") == 0)
                {
                    try
                    {
                        DataTable dtExcel = new DataTable();
                        dtExcel = ReadExcel(txtFilePath.Text, fileExt); //read excel file 

                        if (dtExcel.Rows.Count > 0)
                        {
                            DataTable dtError = new DataTable();
                            dtError.Columns.Add("Contact Name");
                            dtError.Columns.Add("Error Column Name");
                            dtError.Columns.Add("Error Description");

                            lblProgressBar.Visible = true;
                            CreateUpdateContact(dtExcel, dtError);
                            lblProgressBar.Text = "Exporting Error log to excel......";
                            DisplaySummary(dtExcel.Rows.Count, dtError);
                            ExportDataToExcel(dtError);
                            lblProgressBar.Text = string.Format("Import is complete... Log file is generated at {0}", errorFilePath);

                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("File is not in Excel Format");
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            txtFilePath.Text = "";
            rdBtnCreateContact.Checked = false;
            rdBtnUpdateContact.Checked = false;
        }

        private void lnkLblHelp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                VisitLink();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to open link that was clicked.");
            }
        }
        private void VisitLink()
        {
            // Change the color of the link text by setting LinkVisited
            // to true.
            lnkLblHelp.LinkVisited = true;
            //Call the Process.Start method to open the default browser
            //with a URL:
            System.Diagnostics.Process.Start("https://teams.microsoft.com/l/file/5E1E45E0-DF71-4C12-A43A-23C8AA440D36?tenantId=4ae48b41-0137-4599-8661-fc641fe77bea&fileType=docx&objectUrl=https%3A%2F%2Farup.sharepoint.com%2Fsites%2FClientSystemsTeam%2FShared%20Documents%2FInfrastructure%2FTest.docx&baseUrl=https%3A%2F%2Farup.sharepoint.com%2Fsites%2FClientSystemsTeam&serviceName=teams&threadId=19:9fb7082194a14588a73424e55bb81eab@thread.skype&groupId=735cff3c-a8fa-43e5-a965-29c9d33e7d3f");
        }
    }
    public static class FieldConstant
    {
        public static int OPTIONSET_ROLE_VALUE_DECISIONMAKER = 1;
        public static int OPTIONSET_ROLE_VALUE_EMPLOYEE = 2;
        public static int OPTIONSET_ROLE_VALUE_INFLUENCER = 3;
        public static int OPTIONSET_ROLE_VALUE_CSUITE = 770000000;

        public static string OPTIONSET_ROLE_KEY_DECISIONMAKER = "Decision Maker";
        public static string OPTIONSET_ROLE_KEY_EMPLOYEE = "Employee";
        public static string OPTIONSET_ROLE_KEY_INFLUENCER = "Influencer";
        public static string OPTIONSET_ROLE_KEY_CSUITE = "C-Suite";

        public static string COUNTRY_AUSTRALIA = "AUSTRALIA";
        public static string COUNTRY_CANADA = "CANADA";
        public static string COUNTRY_INDONESIA = "INDONESIA";
        public static string COUNTRY_NEW_ZEALAND = "NEW ZEALAND";
        public static string COUNTRY_SINGAPORE = "SINGAPORE";
        public static string COUNTRY_USA = "UNITED STATES OF AMERICA";
        public static string COUNTRY_MALAYSIA = "MALAYSIA";

        public static string OPTIONSET_CONTACTTYPE_KEY_MARKETING = "Marketing Contact";
        public static string OPTIONSET_CONTACTTYPE_KEY_CLIENT = "Client Contact";

        public static int OPTIONSET_CONTACTTYPE_VALUE_MARKETING = 770000001;
        public static int OPTIONSET_CONTACTTYPE_VALUE_CLIENT = 770000000;
   

    }



}