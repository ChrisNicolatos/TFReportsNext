<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmTFReports
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.trvReportSelector = New System.Windows.Forms.TreeView()
        Me.lstBSPFortnight = New System.Windows.Forms.ListBox()
        Me.lblBSPFortnight = New System.Windows.Forms.Label()
        Me.fraCustomer = New System.Windows.Forms.GroupBox()
        Me.optAllCustomers = New System.Windows.Forms.RadioButton()
        Me.optCustomerGroup = New System.Windows.Forms.RadioButton()
        Me.optCustomer = New System.Windows.Forms.RadioButton()
        Me.cmbCustomerGroup = New System.Windows.Forms.ComboBox()
        Me.cmbCustomers = New System.Windows.Forms.ComboBox()
        Me.lstBSPMonth = New System.Windows.Forms.ListBox()
        Me.lblBSPMonth = New System.Windows.Forms.Label()
        Me.opt02Domestic = New System.Windows.Forms.RadioButton()
        Me.opt02International = New System.Windows.Forms.RadioButton()
        Me.chkCheckOption1 = New System.Windows.Forms.CheckBox()
        Me.lstYear = New System.Windows.Forms.ListBox()
        Me.lstMonth = New System.Windows.Forms.ListBox()
        Me.lblMonth = New System.Windows.Forms.Label()
        Me.lblYear = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblDate1 = New System.Windows.Forms.Label()
        Me.lblDate2 = New System.Windows.Forms.Label()
        Me.dtpDate2From = New System.Windows.Forms.DateTimePicker()
        Me.dtpDate2To = New System.Windows.Forms.DateTimePicker()
        Me.savExcel = New System.Windows.Forms.SaveFileDialog()
        Me.dtpDate1To = New System.Windows.Forms.DateTimePicker()
        Me.lblDate1To = New System.Windows.Forms.Label()
        Me.fraReportOptions = New System.Windows.Forms.GroupBox()
        Me.lstGroupList = New System.Windows.Forms.ListBox()
        Me.fraOptionsTriplet = New System.Windows.Forms.GroupBox()
        Me.optTriplet2 = New System.Windows.Forms.RadioButton()
        Me.optTriplet1 = New System.Windows.Forms.RadioButton()
        Me.optTriplet0 = New System.Windows.Forms.RadioButton()
        Me.lblGroupList = New System.Windows.Forms.Label()
        Me.txtTextEntry = New System.Windows.Forms.TextBox()
        Me.lblTextEntry = New System.Windows.Forms.Label()
        Me.lblDate2To = New System.Windows.Forms.Label()
        Me.dtpDate1From = New System.Windows.Forms.DateTimePicker()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.lblVersion = New System.Windows.Forms.ToolStripStatusLabel()
        Me.cmdExportExcel = New System.Windows.Forms.Button()
        Me.pbarCounter = New System.Windows.Forms.ProgressBar()
        Me.fraCustomer.SuspendLayout()
        Me.fraReportOptions.SuspendLayout()
        Me.fraOptionsTriplet.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'trvReportSelector
        '
        Me.trvReportSelector.Location = New System.Drawing.Point(12, 41)
        Me.trvReportSelector.Name = "trvReportSelector"
        Me.trvReportSelector.Size = New System.Drawing.Size(288, 451)
        Me.trvReportSelector.TabIndex = 30
        '
        'lstBSPFortnight
        '
        Me.lstBSPFortnight.FormattingEnabled = True
        Me.lstBSPFortnight.Location = New System.Drawing.Point(460, 163)
        Me.lstBSPFortnight.Name = "lstBSPFortnight"
        Me.lstBSPFortnight.Size = New System.Drawing.Size(123, 43)
        Me.lstBSPFortnight.TabIndex = 7
        '
        'lblBSPFortnight
        '
        Me.lblBSPFortnight.AutoSize = True
        Me.lblBSPFortnight.Location = New System.Drawing.Point(460, 148)
        Me.lblBSPFortnight.Name = "lblBSPFortnight"
        Me.lblBSPFortnight.Size = New System.Drawing.Size(72, 13)
        Me.lblBSPFortnight.TabIndex = 6
        Me.lblBSPFortnight.Text = "BSP Fortnight"
        '
        'fraCustomer
        '
        Me.fraCustomer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraCustomer.Controls.Add(Me.optAllCustomers)
        Me.fraCustomer.Controls.Add(Me.optCustomerGroup)
        Me.fraCustomer.Controls.Add(Me.optCustomer)
        Me.fraCustomer.Controls.Add(Me.cmbCustomerGroup)
        Me.fraCustomer.Controls.Add(Me.cmbCustomers)
        Me.fraCustomer.Location = New System.Drawing.Point(308, 41)
        Me.fraCustomer.Name = "fraCustomer"
        Me.fraCustomer.Size = New System.Drawing.Size(751, 99)
        Me.fraCustomer.TabIndex = 26
        Me.fraCustomer.TabStop = False
        Me.fraCustomer.Text = "Customer"
        Me.fraCustomer.Visible = False
        '
        'optAllCustomers
        '
        Me.optAllCustomers.AutoSize = True
        Me.optAllCustomers.Location = New System.Drawing.Point(124, 19)
        Me.optAllCustomers.Name = "optAllCustomers"
        Me.optAllCustomers.Size = New System.Drawing.Size(87, 17)
        Me.optAllCustomers.TabIndex = 18
        Me.optAllCustomers.TabStop = True
        Me.optAllCustomers.Text = "All customers"
        Me.optAllCustomers.UseVisualStyleBackColor = True
        '
        'optCustomerGroup
        '
        Me.optCustomerGroup.AutoSize = True
        Me.optCustomerGroup.Enabled = False
        Me.optCustomerGroup.Location = New System.Drawing.Point(6, 69)
        Me.optCustomerGroup.Name = "optCustomerGroup"
        Me.optCustomerGroup.Size = New System.Drawing.Size(101, 17)
        Me.optCustomerGroup.TabIndex = 17
        Me.optCustomerGroup.TabStop = True
        Me.optCustomerGroup.Text = "Customer Group"
        Me.optCustomerGroup.UseVisualStyleBackColor = True
        '
        'optCustomer
        '
        Me.optCustomer.AutoSize = True
        Me.optCustomer.Enabled = False
        Me.optCustomer.Location = New System.Drawing.Point(6, 44)
        Me.optCustomer.Name = "optCustomer"
        Me.optCustomer.Size = New System.Drawing.Size(69, 17)
        Me.optCustomer.TabIndex = 16
        Me.optCustomer.TabStop = True
        Me.optCustomer.Text = "Customer"
        Me.optCustomer.UseVisualStyleBackColor = True
        '
        'cmbCustomerGroup
        '
        Me.cmbCustomerGroup.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbCustomerGroup.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbCustomerGroup.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbCustomerGroup.FormattingEnabled = True
        Me.cmbCustomerGroup.Location = New System.Drawing.Point(124, 67)
        Me.cmbCustomerGroup.Name = "cmbCustomerGroup"
        Me.cmbCustomerGroup.Size = New System.Drawing.Size(615, 21)
        Me.cmbCustomerGroup.TabIndex = 9
        '
        'cmbCustomers
        '
        Me.cmbCustomers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbCustomers.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbCustomers.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbCustomers.FormattingEnabled = True
        Me.cmbCustomers.Location = New System.Drawing.Point(124, 42)
        Me.cmbCustomers.Name = "cmbCustomers"
        Me.cmbCustomers.Size = New System.Drawing.Size(615, 21)
        Me.cmbCustomers.TabIndex = 8
        '
        'lstBSPMonth
        '
        Me.lstBSPMonth.FormattingEnabled = True
        Me.lstBSPMonth.Location = New System.Drawing.Point(337, 163)
        Me.lstBSPMonth.Name = "lstBSPMonth"
        Me.lstBSPMonth.Size = New System.Drawing.Size(117, 43)
        Me.lstBSPMonth.TabIndex = 3
        '
        'lblBSPMonth
        '
        Me.lblBSPMonth.AutoSize = True
        Me.lblBSPMonth.Location = New System.Drawing.Point(337, 148)
        Me.lblBSPMonth.Name = "lblBSPMonth"
        Me.lblBSPMonth.Size = New System.Drawing.Size(61, 13)
        Me.lblBSPMonth.TabIndex = 2
        Me.lblBSPMonth.Text = "BSP Month"
        '
        'opt02Domestic
        '
        Me.opt02Domestic.AutoSize = True
        Me.opt02Domestic.Location = New System.Drawing.Point(19, 121)
        Me.opt02Domestic.Name = "opt02Domestic"
        Me.opt02Domestic.Size = New System.Drawing.Size(69, 17)
        Me.opt02Domestic.TabIndex = 1
        Me.opt02Domestic.Text = "Domestic"
        Me.opt02Domestic.UseVisualStyleBackColor = True
        '
        'opt02International
        '
        Me.opt02International.AutoSize = True
        Me.opt02International.Checked = True
        Me.opt02International.Location = New System.Drawing.Point(19, 96)
        Me.opt02International.Name = "opt02International"
        Me.opt02International.Size = New System.Drawing.Size(83, 17)
        Me.opt02International.TabIndex = 0
        Me.opt02International.TabStop = True
        Me.opt02International.Text = "International"
        Me.opt02International.UseVisualStyleBackColor = True
        '
        'chkCheckOption1
        '
        Me.chkCheckOption1.AutoSize = True
        Me.chkCheckOption1.Location = New System.Drawing.Point(19, 71)
        Me.chkCheckOption1.Name = "chkCheckOption1"
        Me.chkCheckOption1.Size = New System.Drawing.Size(151, 17)
        Me.chkCheckOption1.TabIndex = 4
        Me.chkCheckOption1.Text = "Differences as percentage"
        Me.chkCheckOption1.UseVisualStyleBackColor = True
        '
        'lstYear
        '
        Me.lstYear.FormattingEnabled = True
        Me.lstYear.Location = New System.Drawing.Point(16, 163)
        Me.lstYear.Name = "lstYear"
        Me.lstYear.Size = New System.Drawing.Size(79, 43)
        Me.lstYear.TabIndex = 3
        '
        'lstMonth
        '
        Me.lstMonth.FormattingEnabled = True
        Me.lstMonth.Location = New System.Drawing.Point(95, 163)
        Me.lstMonth.Name = "lstMonth"
        Me.lstMonth.Size = New System.Drawing.Size(79, 43)
        Me.lstMonth.TabIndex = 2
        '
        'lblMonth
        '
        Me.lblMonth.AutoSize = True
        Me.lblMonth.Location = New System.Drawing.Point(95, 148)
        Me.lblMonth.Name = "lblMonth"
        Me.lblMonth.Size = New System.Drawing.Size(37, 13)
        Me.lblMonth.TabIndex = 1
        Me.lblMonth.Text = "Month"
        '
        'lblYear
        '
        Me.lblYear.AutoSize = True
        Me.lblYear.Location = New System.Drawing.Point(13, 148)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(29, 13)
        Me.lblYear.TabIndex = 0
        Me.lblYear.Text = "Year"
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.lblTitle.Location = New System.Drawing.Point(308, 6)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(754, 27)
        Me.lblTitle.TabIndex = 27
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDate1
        '
        Me.lblDate1.AutoSize = True
        Me.lblDate1.BackColor = System.Drawing.SystemColors.Control
        Me.lblDate1.Location = New System.Drawing.Point(16, 25)
        Me.lblDate1.Name = "lblDate1"
        Me.lblDate1.Size = New System.Drawing.Size(39, 13)
        Me.lblDate1.TabIndex = 13
        Me.lblDate1.Text = "Date 1"
        '
        'lblDate2
        '
        Me.lblDate2.AutoSize = True
        Me.lblDate2.BackColor = System.Drawing.SystemColors.Control
        Me.lblDate2.Location = New System.Drawing.Point(16, 50)
        Me.lblDate2.Name = "lblDate2"
        Me.lblDate2.Size = New System.Drawing.Size(39, 13)
        Me.lblDate2.TabIndex = 12
        Me.lblDate2.Text = "Date 2"
        '
        'dtpDate2From
        '
        Me.dtpDate2From.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDate2From.Location = New System.Drawing.Point(203, 46)
        Me.dtpDate2From.Name = "dtpDate2From"
        Me.dtpDate2From.Size = New System.Drawing.Size(99, 20)
        Me.dtpDate2From.TabIndex = 8
        '
        'dtpDate2To
        '
        Me.dtpDate2To.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDate2To.Location = New System.Drawing.Point(328, 46)
        Me.dtpDate2To.Name = "dtpDate2To"
        Me.dtpDate2To.Size = New System.Drawing.Size(99, 20)
        Me.dtpDate2To.TabIndex = 9
        '
        'dtpDate1To
        '
        Me.dtpDate1To.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDate1To.Location = New System.Drawing.Point(328, 21)
        Me.dtpDate1To.Name = "dtpDate1To"
        Me.dtpDate1To.Size = New System.Drawing.Size(99, 20)
        Me.dtpDate1To.TabIndex = 5
        '
        'lblDate1To
        '
        Me.lblDate1To.AutoSize = True
        Me.lblDate1To.Location = New System.Drawing.Point(310, 25)
        Me.lblDate1To.Name = "lblDate1To"
        Me.lblDate1To.Size = New System.Drawing.Size(10, 13)
        Me.lblDate1To.TabIndex = 7
        Me.lblDate1To.Text = "-"
        '
        'fraReportOptions
        '
        Me.fraReportOptions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraReportOptions.Controls.Add(Me.lstGroupList)
        Me.fraReportOptions.Controls.Add(Me.fraOptionsTriplet)
        Me.fraReportOptions.Controls.Add(Me.lblGroupList)
        Me.fraReportOptions.Controls.Add(Me.lblBSPFortnight)
        Me.fraReportOptions.Controls.Add(Me.lstBSPFortnight)
        Me.fraReportOptions.Controls.Add(Me.lstBSPMonth)
        Me.fraReportOptions.Controls.Add(Me.txtTextEntry)
        Me.fraReportOptions.Controls.Add(Me.lblBSPMonth)
        Me.fraReportOptions.Controls.Add(Me.lblTextEntry)
        Me.fraReportOptions.Controls.Add(Me.opt02Domestic)
        Me.fraReportOptions.Controls.Add(Me.chkCheckOption1)
        Me.fraReportOptions.Controls.Add(Me.opt02International)
        Me.fraReportOptions.Controls.Add(Me.lblDate2To)
        Me.fraReportOptions.Controls.Add(Me.lstYear)
        Me.fraReportOptions.Controls.Add(Me.lblDate1)
        Me.fraReportOptions.Controls.Add(Me.lstMonth)
        Me.fraReportOptions.Controls.Add(Me.lblDate2)
        Me.fraReportOptions.Controls.Add(Me.lblMonth)
        Me.fraReportOptions.Controls.Add(Me.lblYear)
        Me.fraReportOptions.Controls.Add(Me.dtpDate2From)
        Me.fraReportOptions.Controls.Add(Me.dtpDate2To)
        Me.fraReportOptions.Controls.Add(Me.dtpDate1From)
        Me.fraReportOptions.Controls.Add(Me.dtpDate1To)
        Me.fraReportOptions.Controls.Add(Me.lblDate1To)
        Me.fraReportOptions.Location = New System.Drawing.Point(308, 146)
        Me.fraReportOptions.Name = "fraReportOptions"
        Me.fraReportOptions.Size = New System.Drawing.Size(751, 346)
        Me.fraReportOptions.TabIndex = 25
        Me.fraReportOptions.TabStop = False
        Me.fraReportOptions.Text = "Report Options"
        Me.fraReportOptions.Visible = False
        '
        'lstGroupList
        '
        Me.lstGroupList.FormattingEnabled = True
        Me.lstGroupList.Location = New System.Drawing.Point(534, 43)
        Me.lstGroupList.Name = "lstGroupList"
        Me.lstGroupList.Size = New System.Drawing.Size(205, 95)
        Me.lstGroupList.TabIndex = 20
        '
        'fraOptionsTriplet
        '
        Me.fraOptionsTriplet.Controls.Add(Me.optTriplet2)
        Me.fraOptionsTriplet.Controls.Add(Me.optTriplet1)
        Me.fraOptionsTriplet.Controls.Add(Me.optTriplet0)
        Me.fraOptionsTriplet.Location = New System.Drawing.Point(16, 212)
        Me.fraOptionsTriplet.Name = "fraOptionsTriplet"
        Me.fraOptionsTriplet.Size = New System.Drawing.Size(342, 89)
        Me.fraOptionsTriplet.TabIndex = 19
        Me.fraOptionsTriplet.TabStop = False
        '
        'optTriplet2
        '
        Me.optTriplet2.AutoSize = True
        Me.optTriplet2.Location = New System.Drawing.Point(6, 61)
        Me.optTriplet2.Name = "optTriplet2"
        Me.optTriplet2.Size = New System.Drawing.Size(48, 17)
        Me.optTriplet2.TabIndex = 3
        Me.optTriplet2.TabStop = True
        Me.optTriplet2.Text = "Opt2"
        Me.optTriplet2.UseVisualStyleBackColor = True
        '
        'optTriplet1
        '
        Me.optTriplet1.AutoSize = True
        Me.optTriplet1.Location = New System.Drawing.Point(6, 38)
        Me.optTriplet1.Name = "optTriplet1"
        Me.optTriplet1.Size = New System.Drawing.Size(48, 17)
        Me.optTriplet1.TabIndex = 2
        Me.optTriplet1.TabStop = True
        Me.optTriplet1.Text = "Opt1"
        Me.optTriplet1.UseVisualStyleBackColor = True
        '
        'optTriplet0
        '
        Me.optTriplet0.AutoSize = True
        Me.optTriplet0.Location = New System.Drawing.Point(6, 15)
        Me.optTriplet0.Name = "optTriplet0"
        Me.optTriplet0.Size = New System.Drawing.Size(48, 17)
        Me.optTriplet0.TabIndex = 1
        Me.optTriplet0.TabStop = True
        Me.optTriplet0.Text = "Opt0"
        Me.optTriplet0.UseVisualStyleBackColor = True
        '
        'lblGroupList
        '
        Me.lblGroupList.AutoSize = True
        Me.lblGroupList.Location = New System.Drawing.Point(531, 26)
        Me.lblGroupList.Name = "lblGroupList"
        Me.lblGroupList.Size = New System.Drawing.Size(87, 13)
        Me.lblGroupList.TabIndex = 17
        Me.lblGroupList.Text = "OperationsGroup"
        '
        'txtTextEntry
        '
        Me.txtTextEntry.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTextEntry.Location = New System.Drawing.Point(186, 163)
        Me.txtTextEntry.Multiline = True
        Me.txtTextEntry.Name = "txtTextEntry"
        Me.txtTextEntry.Size = New System.Drawing.Size(143, 43)
        Me.txtTextEntry.TabIndex = 1
        '
        'lblTextEntry
        '
        Me.lblTextEntry.AutoSize = True
        Me.lblTextEntry.Location = New System.Drawing.Point(186, 148)
        Me.lblTextEntry.Name = "lblTextEntry"
        Me.lblTextEntry.Size = New System.Drawing.Size(146, 13)
        Me.lblTextEntry.TabIndex = 0
        Me.lblTextEntry.Text = "Ticket Numbers (one per line)"
        '
        'lblDate2To
        '
        Me.lblDate2To.AutoSize = True
        Me.lblDate2To.Location = New System.Drawing.Point(310, 50)
        Me.lblDate2To.Name = "lblDate2To"
        Me.lblDate2To.Size = New System.Drawing.Size(10, 13)
        Me.lblDate2To.TabIndex = 14
        Me.lblDate2To.Text = "-"
        '
        'dtpDate1From
        '
        Me.dtpDate1From.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDate1From.Location = New System.Drawing.Point(203, 21)
        Me.dtpDate1From.Name = "dtpDate1From"
        Me.dtpDate1From.Size = New System.Drawing.Size(99, 20)
        Me.dtpDate1From.TabIndex = 4
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblVersion})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 482)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1071, 22)
        Me.StatusStrip1.TabIndex = 31
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'lblVersion
        '
        Me.lblVersion.Name = "lblVersion"
        Me.lblVersion.Size = New System.Drawing.Size(45, 17)
        Me.lblVersion.Text = "Version"
        '
        'cmdExportExcel
        '
        Me.cmdExportExcel.Location = New System.Drawing.Point(8, 6)
        Me.cmdExportExcel.Name = "cmdExportExcel"
        Me.cmdExportExcel.Size = New System.Drawing.Size(108, 27)
        Me.cmdExportExcel.TabIndex = 32
        Me.cmdExportExcel.Text = "Export Excel"
        Me.cmdExportExcel.UseVisualStyleBackColor = True
        '
        'pbarCounter
        '
        Me.pbarCounter.Location = New System.Drawing.Point(237, 6)
        Me.pbarCounter.Name = "pbarCounter"
        Me.pbarCounter.Size = New System.Drawing.Size(63, 27)
        Me.pbarCounter.TabIndex = 33
        Me.pbarCounter.Visible = False
        '
        'FrmTFReports
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1071, 504)
        Me.Controls.Add(Me.pbarCounter)
        Me.Controls.Add(Me.cmdExportExcel)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.fraReportOptions)
        Me.Controls.Add(Me.trvReportSelector)
        Me.Controls.Add(Me.fraCustomer)
        Me.Controls.Add(Me.lblTitle)
        Me.Name = "FrmTFReports"
        Me.Text = "ATPI ATH Reports"
        Me.fraCustomer.ResumeLayout(False)
        Me.fraCustomer.PerformLayout()
        Me.fraReportOptions.ResumeLayout(False)
        Me.fraReportOptions.PerformLayout()
        Me.fraOptionsTriplet.ResumeLayout(False)
        Me.fraOptionsTriplet.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents trvReportSelector As TreeView
    Friend WithEvents lstBSPFortnight As ListBox
    Friend WithEvents lblBSPFortnight As Label
    Friend WithEvents fraCustomer As GroupBox
    Friend WithEvents cmbCustomers As ComboBox
    Friend WithEvents lstBSPMonth As ListBox
    Friend WithEvents lblBSPMonth As Label
    Friend WithEvents opt02Domestic As RadioButton
    Friend WithEvents opt02International As RadioButton
    Friend WithEvents lblTitle As Label
    Friend WithEvents lblDate1 As Label
    Friend WithEvents lblDate2 As Label
    Friend WithEvents dtpDate2From As DateTimePicker
    Friend WithEvents dtpDate2To As DateTimePicker
    Friend WithEvents savExcel As SaveFileDialog
    Friend WithEvents dtpDate1To As DateTimePicker
    Friend WithEvents lblDate1To As Label
    Friend WithEvents fraReportOptions As GroupBox
    Friend WithEvents dtpDate1From As DateTimePicker
    Friend WithEvents lstYear As ListBox
    Friend WithEvents lstMonth As ListBox
    Friend WithEvents lblMonth As Label
    Friend WithEvents lblYear As Label
    Friend WithEvents txtTextEntry As TextBox
    Friend WithEvents lblTextEntry As Label
    Friend WithEvents chkCheckOption1 As CheckBox
    Friend WithEvents lblDate2To As Label
    Friend WithEvents lblGroupList As Label
    Friend WithEvents cmbCustomerGroup As ComboBox
    Friend WithEvents optCustomerGroup As RadioButton
    Friend WithEvents optCustomer As RadioButton
    Friend WithEvents fraOptionsTriplet As GroupBox
    Friend WithEvents optTriplet2 As RadioButton
    Friend WithEvents optTriplet1 As RadioButton
    Friend WithEvents optTriplet0 As RadioButton
    Friend WithEvents optAllCustomers As RadioButton
    Friend WithEvents lstGroupList As ListBox
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents lblVersion As ToolStripStatusLabel
    Friend WithEvents cmdExportExcel As Button
    Friend WithEvents pbarCounter As ProgressBar
End Class
