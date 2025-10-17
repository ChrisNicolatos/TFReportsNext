using System;
namespace TFReportsNext
{
    public partial class FormTFReportsNext
    {
        private void PrepareOptions()
        {
            dtpDate1From.Value = mobjReports.Date1From;
            dtpDate1To.Value = mobjReports.Date1To;
            dtpDate2From.Value = mobjReports.Date2From;
            dtpDate2To.Value = mobjReports.Date2To;

            if (!string.IsNullOrEmpty(mobjReports.SelectedCustomer))
            {
                cmbCustomers.SelectedIndex = cmbCustomers.FindString(mobjReports.SelectedCustomer);
            }
            else if (mobjSelectedReport != null && !string.IsNullOrEmpty(mobjSelectedReport.InitialClientCode))
            {
                mobjReports.SelectedCustomer = mobjSelectedReport.InitialClientCode;
                optCustomer.Checked = true;
                cmbCustomers.SelectedIndex = cmbCustomers.FindString(mobjSelectedReport.InitialClientCode);
            }

            cmbCustomerGroup.SelectedIndex = cmbCustomerGroup.FindString(mobjReports.CustomerGroup);

            // BSP Month Report by airline
            if (mobjReports.MonthDomestic)
            {
                opt02Domestic.Checked = true;
            }
            else
            {
                opt02International.Checked = true;
            }

            //if (lstBSPMonth.Items.Count > 0)
            //{
            //    lstBSPMonth.SelectedIndex = 0;
            //    if (lstBSPMonth.SelectedItem != null)
            //    {
            //        mobjReports.BSPMonthDate = lstBSPMonth.SelectedItem.ToString();
            //    }
            //    else
            //    {
            //        mobjReports.BSPMonthDate = null;
            //    }
            //}
            //if (mobjReports.BSPMonthDate != null && lstBSPMonth.SelectedItem != null)
            //{
            //    mobjReports.BSPMonthDate = lstBSPMonth.SelectedItem.ToString();
            //}

            //// BSP Fortnight Report by ticket
            //if (lstBSPFortnight.Items.Count > 0)
            //{
            //    lstBSPFortnight.SelectedIndex = 0;
            //}

            //if (lstBSPFortnight.Items.Count > 0 && lstBSPFortnight.SelectedItem != null)
            //{
            //    mobjReports.BSPFortDate = Convert.ToDateTime(lstBSPFortnight.SelectedItem);
            //}

            chkCheckOption1.Checked = mobjReports.BooleanOption1;
        }
        private void ValidateOptions()
        {
            if (mobjSelectedReport == null) return;

            bool pflgLoading = mflgLoading;
            mflgLoading = true;

            SetClientTypeOptions();
            SetTripletOptions();

            mflgLoading = pflgLoading;

            cmdExportExcel.Enabled = mobjSelectedReport.Index switch
            {
                //2 => !string.IsNullOrEmpty(mobjReports.BSPMonthDate),
                //3 => mobjReports.BSPFortDate != DateTime.MinValue,
                4 => mobjReports.TextEntryItemsCount > 0,
                6 => lstYear.SelectedIndex >= 0 && lstMonth.SelectedIndex > 0,
                12 => lstYear.SelectedIndex >= 0 && lstMonth.SelectedIndex >= 0,
                17 => CheckClientTypeOption(),
                18 => mobjReports.OptionTriplet >= 0 && mobjReports.OptionTriplet <= 2 && CheckClientTypeOption(),
                19 => CheckClientTypeOption(),
                21 => true,
                _ => (mobjReports.Date1Checked && mobjReports.Date1From <= mobjReports.Date1To) ||
                     (mobjReports.Date2Checked && mobjReports.Date2From <= mobjReports.Date2To)
            };

            if (mobjSelectedReport.Index == 22)
            {
                if (mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.ByClient)
                {
                    chkCheckOption1.Enabled = true;
                }
                else
                {
                    chkCheckOption1.Enabled = false;
                    chkCheckOption1.Checked = false;
                }
            }
        }
        private bool CheckClientTypeOption()
        {
            return mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.AllClients ||
                      (mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.ByClient && !string.IsNullOrEmpty(mobjReports.SelectedCustomer)) ||
                      (mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.ByGroup && mobjReports.TagID != 0);
        }
        private void SetClientTypeOptions()
        {
            optAllCustomers.Checked = mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.AllClients;
            optCustomer.Checked = mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.ByClient;
            optCustomerGroup.Checked = mobjReports.ByClient == ReportsNext.ReportsCollection.ClientReportType.ByGroup;
        }

        private void SetTripletOptions()
        {
            optTriplet0.Checked = mobjReports.OptionTriplet == 0;
            optTriplet1.Checked = mobjReports.OptionTriplet == 1;
            optTriplet2.Checked = mobjReports.OptionTriplet == 2;
        }
    }
}