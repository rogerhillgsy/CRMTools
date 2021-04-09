using DisableUserWPFApplication.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Web.UI.WebControls;
using System.Windows;


namespace DisableUserWPFApplication
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string url;
        private string username;
        private string password;
        public MainWindow()
        {
            InitializeComponent();
            if(Convert.ToBoolean(ConfigurationManager.AppSettings["isAdmin"]))
            checkBox.IsEnabled = true;
            else
                checkBox.IsEnabled = false;
        }
       
        public void LoadMarketingList()
        {
            MarketingBL objbl = new MarketingBL(username, password, url);
            var MarketingList = objbl.getMarketingList() ;
            listBoxMarketingList.ItemsSource = MarketingList;
        }
       
        /// <summary>
        /// User Management Submit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Click(object sender, RoutedEventArgs e)
        {
             url = txtUrl.Text;
            var users = txtUsers.Text.Replace("\n",string.Empty).Replace("\r",string.Empty);
             username = txtUser.Text;
             password = TxtPassword.Password;
            try
            {
                if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(users) && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    CRMHelper _crmHelper = new CRMHelper(username, password, url);
                    foreach (var userid in users.Split(','))
                    {
                        var user = _crmHelper.getUserIdFromCRM(userid);
                        if ((bool)radioButton.IsChecked == true)
                        {
                            var i = _crmHelper.DisableUserIfEnabled(user);

                        }
                        else
                        {
                            var i = _crmHelper.EnableUserIfDisabled(user);
                        }
                    }
                    MessageBox.Show("Task Completed successfully");
                }
                else
                {
                    MessageBox.Show("Please enter all the values");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error occured" + ex.Message);
            }
        }

        private void checkBox_Checked(object sender, RoutedEventArgs e)
        {
                txtUser.IsEnabled = false;
                txtUser.Text = ConfigurationManager.AppSettings["UserId"].ToString();
                TxtPassword.IsEnabled = false;
                TxtPassword.Password = ConfigurationManager.AppSettings["Password"].ToString(); 
        }

        private void checkBox_Unchecked(object sender, RoutedEventArgs e)
        {
            txtUser.IsEnabled = true;
            TxtPassword.IsEnabled = true;
            txtUser.Text = "";
            TxtPassword.Password = "";
        }

        /// <summary>
        /// Marketing submit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Marketing_Button_Click(object sender, RoutedEventArgs e)
        {
            // find checked boxes in ML
            var _marketinglist =(MarketingListModel)listBoxMarketingList.SelectedValue;
            var emailids = txtContactEmails.Text.Replace("\n", string.Empty).Replace("\r", string.Empty);
            if (_marketinglist != null && !string.IsNullOrEmpty(emailids))
            {
                MessageBoxResult result = MessageBox.Show("Do you want to continue? You are going to update Marketing list " + _marketinglist.Name,
                                          "Confirmation",
                                          MessageBoxButton.YesNo,
                                          MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    txtlog.Clear();
                    List<string> arguments = new List<string>();
                    arguments.Add(_marketinglist.GuidValue);
                    arguments.Add(emailids);
                    arguments.Add(_marketinglist.Name);
                    Resulttab.IsSelected = true;

                    BackgroundWorker worker = new BackgroundWorker();
                    worker.WorkerReportsProgress = true;
                    worker.DoWork += worker_DoWork;
                    worker.ProgressChanged += worker_ProgressChanged;
                    worker.RunWorkerCompleted += worker_RunWorkerCompleted;
                    worker.RunWorkerAsync(arguments);


                    // MarketingBL objbl = new MarketingBL(username, password, url);
                    //var MarketingList = AssociatecontactwithMarketinglist(_marketinglist,emailids);
                }
                else {
                    return;
                }

            }
            else
            {
                MessageBox.Show("Please enter all the values");
            }
        }

        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                List<string> arguments = (List<string>)e.Argument;

                MarketingBL objbl = new MarketingBL(username, password, url);
                //CRMHelper _crmHelper = new CRMHelper(username, password, url);
                List<MemberDetails> members = objbl.getMarketingListMembers(arguments[0]);

                if (members != null)
                {
                    var text = arguments[2] + " : Members Count " + members.Count;
                    (sender as BackgroundWorker).ReportProgress(1, text);

                }

                string[] emails = arguments[1].Split(',');
                int i = 1;
                bool IsMember = false;
                foreach (string email in emails)
                {
                    int progresspercentage = Convert.ToInt32(((double)i / emails.Length) * 100);
                    //var text = "Processing emailid : " + email;
                    //(sender as BackgroundWorker).ReportProgress(progresspercentage, text);
                    List<ContactModel> contacts = objbl.getContactDetails(email);
                    if (contacts != null && contacts.Count > 0)
                    {
                        if (contacts.Count == 1)
                        {
                            foreach (var cnt in contacts)
                            {
                                if (string.Equals(cnt.status, "1"))
                                {
                                    IsMember = false;
                                    // check in marketing list
                                    foreach (var mem in members)
                                    {
                                        if (Guid.Equals(mem.Memberid, cnt.ID))
                                        {
                                            IsMember = true;
                                            var textprogress = email + ": " + cnt.Name + " is already present in Marketing list : " + arguments[2];
                                            (sender as BackgroundWorker).ReportProgress(progresspercentage, textprogress);
                                        }
                                    }
                                    if (!IsMember)
                                    {
                                      
                                        try
                                        {
                                            objbl.AddMembersToList(arguments[0], cnt.ID);
                                            var textprogress = email + " OK";
                                            (sender as BackgroundWorker).ReportProgress(progresspercentage, textprogress);
                                        }
                                        catch (Exception exe)
                                        {
                                            var textprogress = "Exception occured while adding " + cnt.Name + "to Marketing list " + arguments[2] + "exception : " + exe.Message;
                                            (sender as BackgroundWorker).ReportProgress(progresspercentage, textprogress);
                                        }


                                    }

                                }
                                else
                                {
                                    var textprogress = cnt.Name + ": is inactive in CRM";
                                    (sender as BackgroundWorker).ReportProgress(progresspercentage, textprogress);
                                }

                            }
                        }
                        else
                        {
                           var textprogress = email + ": Multiple Contacts found";
                            (sender as BackgroundWorker).ReportProgress(progresspercentage, textprogress);
                        }
                    }
                    else
                    {
                        string textprogress = email + ": Contact not found";
                        (sender as BackgroundWorker).ReportProgress(progresspercentage, textprogress);
                    }


                    i++;
                }
            }
            catch(Exception ex)
            {
                string textprogress = "Exception Occured : " + ex.Message;
                (sender as BackgroundWorker).ReportProgress(100, textprogress);
            }
        }
       void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressbar.Value = e.ProgressPercentage;
            txtlog.AppendText(e.UserState.ToString());
            txtlog.AppendText(Environment.NewLine);
           
        }

        void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //   MessageBox.Show("Numbers between 0 and 10000 divisible by 7: " + e.Result);
            txtlog.AppendText ("completed");
            progressbar.Value = 100;
            System.Threading.Thread.Sleep(2000);
            progressbar.Value = 0;
        }
        internal object AssociatecontactwithMarketinglist(string _marketinglist, string emailids, System.Windows.Controls.TextBox txtlogWriter)
        {
            // get marketing list members
            MarketingBL objbl = new MarketingBL(username, password, url);
            //CRMHelper _crmHelper = new CRMHelper(username, password, url);
            List<MemberDetails> members = objbl.getMarketingListMembers(_marketinglist);
            if (members != null)
            {
               // txtlogWriter.AppendText(_marketinglist + " : Members Count " + members.Count);
              //  txtlogWriter.AppendText(Environment.NewLine);
            }

            string[] emails = emailids.Split(',');

            foreach (string email in emails)
            {
                List<ContactModel> contacts = objbl.getContactDetails(email);
                if(contacts!=null && contacts.Count > 0)
                {
                    foreach (var cnt in contacts)
                    {
                       // txtlogWriter.AppendText(email + ": Name :" + cnt.Name);
                      //  txtlogWriter.AppendText(Environment.NewLine);
                    }
                }
            }
            return null;

        }
        private void btnConnectCRM_Click(object sender, RoutedEventArgs e)
        {
             url = txtUrl.Text;
             username = txtUser.Text;
             password = TxtPassword.Password;
            listBoxMarketingList.ItemsSource = null;
            listBoxMarketingList.Items.Clear();
            checkBox.IsEnabled = false;
            var users = txtUsers.Text.Replace("\n", string.Empty).Replace("\r", string.Empty);
            try
            {
                if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    LoadMarketingList();
                    txtUrl.IsEnabled = false;
                    TxtPassword.IsEnabled = false;
                    txtUser.IsEnabled = false;
                }
                else
                {
                    MessageBox.Show("Please enter all the values");
                }
                btnConnectCRM.IsEnabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occured" + ex.Message);
            }
          
        }
    }
}
