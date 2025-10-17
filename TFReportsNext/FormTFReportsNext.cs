using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TFReportsNext
{
    public partial class FormTFReportsNext : Form
    {
        const int LOADYEARSMINYEAR = 2000;
        const int LOADBSPMONTHSMINYEAR = 2013;
        bool mflgLoading = false;
        int pLocX = 16;
        int pLocY = 25;
        int pStepX = 25;
        int pStepY = 25;
        ReportsNext.ReportsCollection mobjReports = new ReportsNext.ReportsCollection();
        ReportsNext.ReportsItem? mobjSelectedReport;
        public FormTFReportsNext()
        {
            InitializeComponent();
        }

        private void FormTFReportsNext_Load(object sender, EventArgs e)
        {
            try
            {
                mflgLoading = true;
                SetEvents();
                lblVersion.Text = VersionText;
                fraCustomer.Visible = false;
                fraReportOptions.Visible = false;
                PopulateReportSelector();
                Frames_PrepareFrames();
                PrepareOptions();
                ValidateOptions();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                mflgLoading = false;
            }
        }
        void SetEvents()
        {
            cmdExportExcel.Click += CmdExportExcel_Click;
            cmbCustomers.SelectedIndexChanged += CmbCustomers_SelectedIndexChanged;
            cmbCustomerGroup.SelectedIndexChanged += CmbCustomerGroup_SelectedIndexChanged;
            dtpDate1From.ValueChanged += DtpDate1From_ValueChanged;
            dtpDate1To.ValueChanged += DtpDate1To_ValueChanged;
            dtpDate2From.ValueChanged += DdtpDate2From_ValueChanged;
            dtpDate2To.ValueChanged += DtpDate2To_ValueChanged;
            trvReportSelector.AfterSelect += TrvReportSelector_AfterSelect;
            lstYear.SelectedIndexChanged += LstYear_SelectedIndexChanged;
            lstMonth.SelectedIndexChanged += LstMonth_SelectedIndexChanged;
            chkCheckOption1.CheckedChanged += ChkCheckOption1_CheckedChanged;
            txtTextEntry.TextChanged += TxtTextEntry_TextChanged;
            optTriplet0.CheckedChanged += OptTriplet0_CheckedChanged;
            optTriplet1.CheckedChanged += OptTriplet1_CheckedChanged;
            optTriplet2.CheckedChanged += OptTriplet2_CheckedChanged;
            optAllCustomers.CheckedChanged += OptAllCustomers_CheckedChanged;
            lstGroupList.SelectedIndexChanged += LstGroupList_SelectedIndexChanged;

            //lstBSPMonth.SelectedIndexChanged += LstBSPMonth_SelectedIndexChanged;
            //lstBSPFortnight.SelectedIndexChanged += LstBSPFortnight_SelectedIndexChanged;
        }
        void PopulateReportSelector()
        {
            trvReportSelector.BeginUpdate();
            trvReportSelector.Nodes.Clear();

            // Group reports by GroupName, defaulting to "Miscellaneous"
            var groupedReports = mobjReports.Values
                .Where(r => !r.Hidden)
                .GroupBy(r => string.IsNullOrWhiteSpace(r.GroupName) ? "Miscellaneous" : r.GroupName);

            foreach (var group in groupedReports)
            {
                var groupNode = trvReportSelector.Nodes.Add(group.Key, group.Key);
                foreach (var report in group)
                {
                    // Ensure ReportName is not null
                    var reportName = string.IsNullOrWhiteSpace(report.ReportName) ? "(Unnamed Report)" : report.ReportName;
                    groupNode.Nodes.Add(report.Index.ToString(), reportName);
                }
            }

            trvReportSelector.Sort();
            trvReportSelector.EndUpdate();
        }
    }
}
