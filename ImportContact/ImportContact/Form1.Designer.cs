namespace ImportContact
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lblNumOfRecordProcessed = new System.Windows.Forms.Label();
            this.lblNumOfDuplicateRecords = new System.Windows.Forms.Label();
            this.lblTotalValidationError = new System.Windows.Forms.Label();
            this.lblTotalRecordsImported = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblProgressBar = new System.Windows.Forms.Label();
            this.rdBtnCreateContact = new System.Windows.Forms.RadioButton();
            this.rdBtnUpdateContact = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.grpBoxSummary = new System.Windows.Forms.GroupBox();
            this.lblEnvironment = new System.Windows.Forms.Label();
            this.cmbEnvironment = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lnkLblHelp = new System.Windows.Forms.LinkLabel();
            this.grpBoxSummary.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnBrowse.Location = new System.Drawing.Point(305, 74);
            this.btnBrowse.Margin = new System.Windows.Forms.Padding(2);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(110, 26);
            this.btnBrowse.TabIndex = 0;
            this.btnBrowse.Text = "Browse";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.Button1_Click);
            // 
            // lblNumOfRecordProcessed
            // 
            this.lblNumOfRecordProcessed.AutoSize = true;
            this.lblNumOfRecordProcessed.Location = new System.Drawing.Point(14, 25);
            this.lblNumOfRecordProcessed.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblNumOfRecordProcessed.Name = "lblNumOfRecordProcessed";
            this.lblNumOfRecordProcessed.Size = new System.Drawing.Size(149, 13);
            this.lblNumOfRecordProcessed.TabIndex = 2;
            this.lblNumOfRecordProcessed.Text = "Total # of Records Processed";
            this.lblNumOfRecordProcessed.Visible = false;
            // 
            // lblNumOfDuplicateRecords
            // 
            this.lblNumOfDuplicateRecords.AutoSize = true;
            this.lblNumOfDuplicateRecords.Location = new System.Drawing.Point(14, 47);
            this.lblNumOfDuplicateRecords.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblNumOfDuplicateRecords.Name = "lblNumOfDuplicateRecords";
            this.lblNumOfDuplicateRecords.Size = new System.Drawing.Size(148, 13);
            this.lblNumOfDuplicateRecords.TabIndex = 3;
            this.lblNumOfDuplicateRecords.Text = "Total # of Possible Duplicates";
            this.lblNumOfDuplicateRecords.Visible = false;
            // 
            // lblTotalValidationError
            // 
            this.lblTotalValidationError.AutoSize = true;
            this.lblTotalValidationError.Location = new System.Drawing.Point(14, 73);
            this.lblTotalValidationError.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblTotalValidationError.Name = "lblTotalValidationError";
            this.lblTotalValidationError.Size = new System.Drawing.Size(132, 13);
            this.lblTotalValidationError.TabIndex = 4;
            this.lblTotalValidationError.Text = "Total # of Validation Errors";
            this.lblTotalValidationError.Visible = false;
            // 
            // lblTotalRecordsImported
            // 
            this.lblTotalRecordsImported.AutoSize = true;
            this.lblTotalRecordsImported.Location = new System.Drawing.Point(14, 98);
            this.lblTotalRecordsImported.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblTotalRecordsImported.Name = "lblTotalRecordsImported";
            this.lblTotalRecordsImported.Size = new System.Drawing.Size(140, 13);
            this.lblTotalRecordsImported.TabIndex = 5;
            this.lblTotalRecordsImported.Text = "Total # of Records Imported";
            this.lblTotalRecordsImported.Visible = false;
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(30, 78);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(2);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(264, 20);
            this.txtFilePath.TabIndex = 5;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(30, 106);
            this.btnOk.Margin = new System.Windows.Forms.Padding(2);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(65, 23);
            this.btnOk.TabIndex = 7;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(119, 106);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(2);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(69, 23);
            this.btnCancel.TabIndex = 8;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(30, 157);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(385, 23);
            this.progressBar1.TabIndex = 10;
            this.progressBar1.Visible = false;
            // 
            // lblProgressBar
            // 
            this.lblProgressBar.AutoSize = true;
            this.lblProgressBar.Location = new System.Drawing.Point(27, 141);
            this.lblProgressBar.Name = "lblProgressBar";
            this.lblProgressBar.Size = new System.Drawing.Size(96, 13);
            this.lblProgressBar.TabIndex = 9;
            this.lblProgressBar.Text = "Progress Bar Label";
            this.lblProgressBar.Visible = false;
            // 
            // rdBtnCreateContact
            // 
            this.rdBtnCreateContact.AutoSize = true;
            this.rdBtnCreateContact.Location = new System.Drawing.Point(134, 49);
            this.rdBtnCreateContact.Name = "rdBtnCreateContact";
            this.rdBtnCreateContact.Size = new System.Drawing.Size(96, 17);
            this.rdBtnCreateContact.TabIndex = 3;
            this.rdBtnCreateContact.TabStop = true;
            this.rdBtnCreateContact.Text = "Create Contact";
            this.rdBtnCreateContact.UseVisualStyleBackColor = true;
            // 
            // rdBtnUpdateContact
            // 
            this.rdBtnUpdateContact.AutoSize = true;
            this.rdBtnUpdateContact.Location = new System.Drawing.Point(262, 47);
            this.rdBtnUpdateContact.Name = "rdBtnUpdateContact";
            this.rdBtnUpdateContact.Size = new System.Drawing.Size(100, 17);
            this.rdBtnUpdateContact.TabIndex = 4;
            this.rdBtnUpdateContact.TabStop = true;
            this.rdBtnUpdateContact.Text = "Update Contact";
            this.rdBtnUpdateContact.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(74, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Request Type";
            // 
            // grpBoxSummary
            // 
            this.grpBoxSummary.Controls.Add(this.lblNumOfRecordProcessed);
            this.grpBoxSummary.Controls.Add(this.lblNumOfDuplicateRecords);
            this.grpBoxSummary.Controls.Add(this.lblTotalValidationError);
            this.grpBoxSummary.Controls.Add(this.lblTotalRecordsImported);
            this.grpBoxSummary.Location = new System.Drawing.Point(76, 223);
            this.grpBoxSummary.Name = "grpBoxSummary";
            this.grpBoxSummary.Size = new System.Drawing.Size(264, 123);
            this.grpBoxSummary.TabIndex = 11;
            this.grpBoxSummary.TabStop = false;
            this.grpBoxSummary.Text = "Summary";
            this.grpBoxSummary.Visible = false;
            // 
            // lblEnvironment
            // 
            this.lblEnvironment.AutoSize = true;
            this.lblEnvironment.Location = new System.Drawing.Point(27, 10);
            this.lblEnvironment.Name = "lblEnvironment";
            this.lblEnvironment.Size = new System.Drawing.Size(66, 13);
            this.lblEnvironment.TabIndex = 0;
            this.lblEnvironment.Text = "Environment";
            // 
            // cmbEnvironment
            // 
            this.cmbEnvironment.FormattingEnabled = true;
            this.cmbEnvironment.Items.AddRange(new object[] {
            "UAT",
            "PRODUCTION"});
            this.cmbEnvironment.Location = new System.Drawing.Point(119, 10);
            this.cmbEnvironment.Name = "cmbEnvironment";
            this.cmbEnvironment.Size = new System.Drawing.Size(121, 21);
            this.cmbEnvironment.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.cmbEnvironment);
            this.panel1.Controls.Add(this.lblEnvironment);
            this.panel1.Controls.Add(this.lblProgressBar);
            this.panel1.Controls.Add(this.rdBtnUpdateContact);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.progressBar1);
            this.panel1.Controls.Add(this.btnOk);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.rdBtnCreateContact);
            this.panel1.Controls.Add(this.txtFilePath);
            this.panel1.Controls.Add(this.btnBrowse);
            this.panel1.Location = new System.Drawing.Point(46, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(492, 205);
            this.panel1.TabIndex = 12;
            // 
            // lnkLblHelp
            // 
            this.lnkLblHelp.AutoSize = true;
            this.lnkLblHelp.Location = new System.Drawing.Point(586, 12);
            this.lnkLblHelp.Name = "lnkLblHelp";
            this.lnkLblHelp.Size = new System.Drawing.Size(29, 13);
            this.lnkLblHelp.TabIndex = 13;
            this.lnkLblHelp.TabStop = true;
            this.lnkLblHelp.Text = "Help";
            this.lnkLblHelp.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkLblHelp_LinkClicked);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(653, 418);
            this.Controls.Add(this.lnkLblHelp);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.grpBoxSummary);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "Form1";
            this.Text = "ImportContact";
            this.grpBoxSummary.ResumeLayout(false);
            this.grpBoxSummary.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label lblNumOfRecordProcessed;
        private System.Windows.Forms.Label lblNumOfDuplicateRecords;
        private System.Windows.Forms.Label lblTotalValidationError;
        private System.Windows.Forms.Label lblTotalRecordsImported;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblProgressBar;
        private System.Windows.Forms.RadioButton rdBtnCreateContact;
        private System.Windows.Forms.RadioButton rdBtnUpdateContact;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox grpBoxSummary;
        private System.Windows.Forms.Label lblEnvironment;
        private System.Windows.Forms.ComboBox cmbEnvironment;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.LinkLabel lnkLblHelp;
    }
}

