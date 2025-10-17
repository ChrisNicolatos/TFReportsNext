namespace TFReportsNext
{
    partial class FormTFReportsNext
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
            cmdExportExcel = new Button();
            trvReportSelector = new TreeView();
            statusStrip1 = new StatusStrip();
            lblVersion = new ToolStripStatusLabel();
            lblTitle = new Label();
            fraCustomer = new GroupBox();
            cmbCustomerGroup = new ComboBox();
            cmbCustomers = new ComboBox();
            optCustomerGroup = new RadioButton();
            optCustomer = new RadioButton();
            optAllCustomers = new RadioButton();
            fraReportOptions = new GroupBox();
            fraOptionsTriplet = new GroupBox();
            optTriplet2 = new RadioButton();
            optTriplet1 = new RadioButton();
            optTriplet0 = new RadioButton();
            txtTextEntry = new TextBox();
            lblTextEntry = new Label();
            lstGroupList = new ListBox();
            lblGroupList = new Label();
            lstMonth = new ListBox();
            lblMonth = new Label();
            lstYear = new ListBox();
            lblYear = new Label();
            opt02Domestic = new RadioButton();
            opt02International = new RadioButton();
            chkCheckOption1 = new CheckBox();
            lblDate2To = new Label();
            lblDate1To = new Label();
            dtpDate2To = new DateTimePicker();
            dtpDate1To = new DateTimePicker();
            dtpDate2From = new DateTimePicker();
            dtpDate1From = new DateTimePicker();
            lblDate1 = new Label();
            lblDate2 = new Label();
            SaveFileDialog = new SaveFileDialog();
            statusStrip1.SuspendLayout();
            fraCustomer.SuspendLayout();
            fraReportOptions.SuspendLayout();
            fraOptionsTriplet.SuspendLayout();
            SuspendLayout();
            // 
            // cmdExportExcel
            // 
            cmdExportExcel.Location = new Point(8, 6);
            cmdExportExcel.Name = "cmdExportExcel";
            cmdExportExcel.Size = new Size(108, 27);
            cmdExportExcel.TabIndex = 0;
            cmdExportExcel.Text = "Export Excel";
            cmdExportExcel.UseVisualStyleBackColor = true;
            // 
            // trvReportSelector
            // 
            trvReportSelector.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            trvReportSelector.Location = new Point(12, 41);
            trvReportSelector.Name = "trvReportSelector";
            trvReportSelector.Size = new Size(288, 438);
            trvReportSelector.TabIndex = 1;
            // 
            // statusStrip1
            // 
            statusStrip1.Items.AddRange(new ToolStripItem[] { lblVersion });
            statusStrip1.Location = new Point(0, 482);
            statusStrip1.Name = "statusStrip1";
            statusStrip1.Size = new Size(1071, 22);
            statusStrip1.TabIndex = 2;
            statusStrip1.Text = "statusStrip1";
            // 
            // lblVersion
            // 
            lblVersion.Name = "lblVersion";
            lblVersion.Size = new Size(45, 17);
            lblVersion.Text = "Version";
            // 
            // lblTitle
            // 
            lblTitle.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            lblTitle.BackColor = Color.Yellow;
            lblTitle.Location = new Point(308, 6);
            lblTitle.Name = "lblTitle";
            lblTitle.Size = new Size(754, 27);
            lblTitle.TabIndex = 3;
            lblTitle.Text = "label1";
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // fraCustomer
            // 
            fraCustomer.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            fraCustomer.Controls.Add(cmbCustomerGroup);
            fraCustomer.Controls.Add(cmbCustomers);
            fraCustomer.Controls.Add(optCustomerGroup);
            fraCustomer.Controls.Add(optCustomer);
            fraCustomer.Controls.Add(optAllCustomers);
            fraCustomer.Location = new Point(308, 41);
            fraCustomer.Name = "fraCustomer";
            fraCustomer.Size = new Size(751, 99);
            fraCustomer.TabIndex = 4;
            fraCustomer.TabStop = false;
            fraCustomer.Text = "Customer";
            // 
            // cmbCustomerGroup
            // 
            cmbCustomerGroup.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            cmbCustomerGroup.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmbCustomerGroup.AutoCompleteSource = AutoCompleteSource.ListItems;
            cmbCustomerGroup.FormattingEnabled = true;
            cmbCustomerGroup.Location = new Point(137, 70);
            cmbCustomerGroup.Name = "cmbCustomerGroup";
            cmbCustomerGroup.Size = new Size(608, 23);
            cmbCustomerGroup.TabIndex = 4;
            // 
            // cmbCustomers
            // 
            cmbCustomers.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            cmbCustomers.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmbCustomers.AutoCompleteSource = AutoCompleteSource.ListItems;
            cmbCustomers.FormattingEnabled = true;
            cmbCustomers.Location = new Point(137, 43);
            cmbCustomers.Name = "cmbCustomers";
            cmbCustomers.Size = new Size(608, 23);
            cmbCustomers.TabIndex = 3;
            // 
            // optCustomerGroup
            // 
            optCustomerGroup.AutoSize = true;
            optCustomerGroup.Location = new Point(11, 72);
            optCustomerGroup.Name = "optCustomerGroup";
            optCustomerGroup.Size = new Size(113, 19);
            optCustomerGroup.TabIndex = 2;
            optCustomerGroup.TabStop = true;
            optCustomerGroup.Text = "Customer Group";
            optCustomerGroup.UseVisualStyleBackColor = true;
            // 
            // optCustomer
            // 
            optCustomer.AutoSize = true;
            optCustomer.Location = new Point(11, 45);
            optCustomer.Name = "optCustomer";
            optCustomer.Size = new Size(77, 19);
            optCustomer.TabIndex = 1;
            optCustomer.TabStop = true;
            optCustomer.Text = "Customer";
            optCustomer.UseVisualStyleBackColor = true;
            // 
            // optAllCustomers
            // 
            optAllCustomers.AutoSize = true;
            optAllCustomers.Location = new Point(11, 18);
            optAllCustomers.Name = "optAllCustomers";
            optAllCustomers.Size = new Size(99, 19);
            optAllCustomers.TabIndex = 0;
            optAllCustomers.TabStop = true;
            optAllCustomers.Text = "All Customers";
            optAllCustomers.UseVisualStyleBackColor = true;
            // 
            // fraReportOptions
            // 
            fraReportOptions.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            fraReportOptions.Controls.Add(fraOptionsTriplet);
            fraReportOptions.Controls.Add(txtTextEntry);
            fraReportOptions.Controls.Add(lblTextEntry);
            fraReportOptions.Controls.Add(lstGroupList);
            fraReportOptions.Controls.Add(lblGroupList);
            fraReportOptions.Controls.Add(lstMonth);
            fraReportOptions.Controls.Add(lblMonth);
            fraReportOptions.Controls.Add(lstYear);
            fraReportOptions.Controls.Add(lblYear);
            fraReportOptions.Controls.Add(opt02Domestic);
            fraReportOptions.Controls.Add(opt02International);
            fraReportOptions.Controls.Add(chkCheckOption1);
            fraReportOptions.Controls.Add(lblDate2To);
            fraReportOptions.Controls.Add(lblDate1To);
            fraReportOptions.Controls.Add(dtpDate2To);
            fraReportOptions.Controls.Add(dtpDate1To);
            fraReportOptions.Controls.Add(dtpDate2From);
            fraReportOptions.Controls.Add(dtpDate1From);
            fraReportOptions.Controls.Add(lblDate1);
            fraReportOptions.Controls.Add(lblDate2);
            fraReportOptions.Location = new Point(308, 140);
            fraReportOptions.Name = "fraReportOptions";
            fraReportOptions.Size = new Size(751, 339);
            fraReportOptions.TabIndex = 5;
            fraReportOptions.TabStop = false;
            fraReportOptions.Text = "Report Options";
            // 
            // fraOptionsTriplet
            // 
            fraOptionsTriplet.Controls.Add(optTriplet2);
            fraOptionsTriplet.Controls.Add(optTriplet1);
            fraOptionsTriplet.Controls.Add(optTriplet0);
            fraOptionsTriplet.Location = new Point(22, 233);
            fraOptionsTriplet.Name = "fraOptionsTriplet";
            fraOptionsTriplet.Size = new Size(577, 98);
            fraOptionsTriplet.TabIndex = 23;
            fraOptionsTriplet.TabStop = false;
            // 
            // optTriplet2
            // 
            optTriplet2.AutoSize = true;
            optTriplet2.Location = new Point(13, 69);
            optTriplet2.Name = "optTriplet2";
            optTriplet2.Size = new Size(51, 19);
            optTriplet2.TabIndex = 2;
            optTriplet2.TabStop = true;
            optTriplet2.Text = "Opt2";
            optTriplet2.UseVisualStyleBackColor = true;
            // 
            // optTriplet1
            // 
            optTriplet1.AutoSize = true;
            optTriplet1.Location = new Point(13, 42);
            optTriplet1.Name = "optTriplet1";
            optTriplet1.Size = new Size(51, 19);
            optTriplet1.TabIndex = 1;
            optTriplet1.TabStop = true;
            optTriplet1.Text = "Opt1";
            optTriplet1.UseVisualStyleBackColor = true;
            // 
            // optTriplet0
            // 
            optTriplet0.AutoSize = true;
            optTriplet0.Location = new Point(13, 15);
            optTriplet0.Name = "optTriplet0";
            optTriplet0.Size = new Size(51, 19);
            optTriplet0.TabIndex = 0;
            optTriplet0.TabStop = true;
            optTriplet0.Text = "Opt0";
            optTriplet0.UseVisualStyleBackColor = true;
            // 
            // txtTextEntry
            // 
            txtTextEntry.Location = new Point(444, 165);
            txtTextEntry.Multiline = true;
            txtTextEntry.Name = "txtTextEntry";
            txtTextEntry.Size = new Size(155, 64);
            txtTextEntry.TabIndex = 22;
            // 
            // lblTextEntry
            // 
            lblTextEntry.AutoSize = true;
            lblTextEntry.Location = new Point(444, 147);
            lblTextEntry.Name = "lblTextEntry";
            lblTextEntry.Size = new Size(164, 15);
            lblTextEntry.TabIndex = 21;
            lblTextEntry.Text = "Ticket Numbers (one per line)";
            // 
            // lstGroupList
            // 
            lstGroupList.FormattingEnabled = true;
            lstGroupList.ItemHeight = 15;
            lstGroupList.Location = new Point(334, 165);
            lstGroupList.Name = "lstGroupList";
            lstGroupList.Size = new Size(72, 64);
            lstGroupList.TabIndex = 20;
            // 
            // lblGroupList
            // 
            lblGroupList.AutoSize = true;
            lblGroupList.Location = new Point(334, 147);
            lblGroupList.Name = "lblGroupList";
            lblGroupList.Size = new Size(101, 15);
            lblGroupList.TabIndex = 19;
            lblGroupList.Text = "Operations Group";
            // 
            // lstMonth
            // 
            lstMonth.FormattingEnabled = true;
            lstMonth.ItemHeight = 15;
            lstMonth.Location = new Point(100, 165);
            lstMonth.Name = "lstMonth";
            lstMonth.Size = new Size(72, 64);
            lstMonth.TabIndex = 14;
            // 
            // lblMonth
            // 
            lblMonth.AutoSize = true;
            lblMonth.Location = new Point(100, 147);
            lblMonth.Name = "lblMonth";
            lblMonth.Size = new Size(43, 15);
            lblMonth.TabIndex = 13;
            lblMonth.Text = "Month";
            // 
            // lstYear
            // 
            lstYear.FormattingEnabled = true;
            lstYear.ItemHeight = 15;
            lstYear.Location = new Point(22, 165);
            lstYear.Name = "lstYear";
            lstYear.Size = new Size(72, 64);
            lstYear.TabIndex = 12;
            // 
            // lblYear
            // 
            lblYear.AutoSize = true;
            lblYear.Location = new Point(22, 147);
            lblYear.Name = "lblYear";
            lblYear.Size = new Size(29, 15);
            lblYear.TabIndex = 11;
            lblYear.Text = "Year";
            // 
            // opt02Domestic
            // 
            opt02Domestic.AutoSize = true;
            opt02Domestic.Location = new Point(19, 123);
            opt02Domestic.Name = "opt02Domestic";
            opt02Domestic.Size = new Size(75, 19);
            opt02Domestic.TabIndex = 10;
            opt02Domestic.TabStop = true;
            opt02Domestic.Text = "Domestic";
            opt02Domestic.UseVisualStyleBackColor = true;
            // 
            // opt02International
            // 
            opt02International.AutoSize = true;
            opt02International.Location = new Point(19, 96);
            opt02International.Name = "opt02International";
            opt02International.Size = new Size(92, 19);
            opt02International.TabIndex = 9;
            opt02International.TabStop = true;
            opt02International.Text = "International";
            opt02International.UseVisualStyleBackColor = true;
            // 
            // chkCheckOption1
            // 
            chkCheckOption1.AutoSize = true;
            chkCheckOption1.Location = new Point(19, 71);
            chkCheckOption1.Name = "chkCheckOption1";
            chkCheckOption1.Size = new Size(82, 19);
            chkCheckOption1.TabIndex = 8;
            chkCheckOption1.Text = "checkBox1";
            chkCheckOption1.UseVisualStyleBackColor = true;
            // 
            // lblDate2To
            // 
            lblDate2To.AutoSize = true;
            lblDate2To.Location = new Point(309, 50);
            lblDate2To.Name = "lblDate2To";
            lblDate2To.Size = new Size(12, 15);
            lblDate2To.TabIndex = 7;
            lblDate2To.Text = "-";
            // 
            // lblDate1To
            // 
            lblDate1To.AutoSize = true;
            lblDate1To.Location = new Point(309, 25);
            lblDate1To.Name = "lblDate1To";
            lblDate1To.Size = new Size(12, 15);
            lblDate1To.TabIndex = 6;
            lblDate1To.Text = "-";
            // 
            // dtpDate2To
            // 
            dtpDate2To.Format = DateTimePickerFormat.Short;
            dtpDate2To.Location = new Point(327, 46);
            dtpDate2To.Name = "dtpDate2To";
            dtpDate2To.Size = new Size(99, 23);
            dtpDate2To.TabIndex = 5;
            // 
            // dtpDate1To
            // 
            dtpDate1To.Format = DateTimePickerFormat.Short;
            dtpDate1To.Location = new Point(327, 21);
            dtpDate1To.Name = "dtpDate1To";
            dtpDate1To.Size = new Size(99, 23);
            dtpDate1To.TabIndex = 4;
            // 
            // dtpDate2From
            // 
            dtpDate2From.Format = DateTimePickerFormat.Short;
            dtpDate2From.Location = new Point(203, 46);
            dtpDate2From.Name = "dtpDate2From";
            dtpDate2From.Size = new Size(99, 23);
            dtpDate2From.TabIndex = 3;
            // 
            // dtpDate1From
            // 
            dtpDate1From.Format = DateTimePickerFormat.Short;
            dtpDate1From.Location = new Point(203, 21);
            dtpDate1From.Name = "dtpDate1From";
            dtpDate1From.Size = new Size(99, 23);
            dtpDate1From.TabIndex = 2;
            // 
            // lblDate1
            // 
            lblDate1.AutoSize = true;
            lblDate1.Location = new Point(16, 25);
            lblDate1.Name = "lblDate1";
            lblDate1.Size = new Size(40, 15);
            lblDate1.TabIndex = 1;
            lblDate1.Text = "Date 1";
            // 
            // lblDate2
            // 
            lblDate2.AutoSize = true;
            lblDate2.Location = new Point(16, 50);
            lblDate2.Name = "lblDate2";
            lblDate2.Size = new Size(40, 15);
            lblDate2.TabIndex = 0;
            lblDate2.Text = "Date 2";
            // 
            // FormTFReportsNext
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1071, 504);
            Controls.Add(fraReportOptions);
            Controls.Add(fraCustomer);
            Controls.Add(lblTitle);
            Controls.Add(statusStrip1);
            Controls.Add(trvReportSelector);
            Controls.Add(cmdExportExcel);
            Name = "FormTFReportsNext";
            Text = "ATPI ATH Reports";
            Load += FormTFReportsNext_Load;
            statusStrip1.ResumeLayout(false);
            statusStrip1.PerformLayout();
            fraCustomer.ResumeLayout(false);
            fraCustomer.PerformLayout();
            fraReportOptions.ResumeLayout(false);
            fraReportOptions.PerformLayout();
            fraOptionsTriplet.ResumeLayout(false);
            fraOptionsTriplet.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button cmdExportExcel;
        private TreeView trvReportSelector;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel lblVersion;
        private Label lblTitle;
        private GroupBox fraCustomer;
        private RadioButton optCustomerGroup;
        private RadioButton optCustomer;
        private RadioButton optAllCustomers;
        private ComboBox cmbCustomerGroup;
        private ComboBox cmbCustomers;
        private GroupBox fraReportOptions;
        private Label lblDate1;
        private Label lblDate2;
        private DateTimePicker dtpDate1From;
        private DateTimePicker dtpDate2To;
        private DateTimePicker dtpDate1To;
        private DateTimePicker dtpDate2From;
        private Label lblDate1To;
        private Label lblDate2To;
        private CheckBox chkCheckOption1;
        private RadioButton opt02Domestic;
        private RadioButton opt02International;
        private Label lblYear;
        private ListBox lstYear;
        private ListBox lstGroupList;
        private Label lblGroupList;
        private ListBox lstMonth;
        private Label lblMonth;
        private Label lblTextEntry;
        private TextBox txtTextEntry;
        private GroupBox fraOptionsTriplet;
        private RadioButton optTriplet2;
        private RadioButton optTriplet1;
        private RadioButton optTriplet0;
        private SaveFileDialog SaveFileDialog;
    }
}