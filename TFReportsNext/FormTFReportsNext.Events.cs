namespace TFReportsNext
{
    public partial class FormTFReportsNext
    {
        System.Data.DataSet ds;

        #region EXPORT AND CLOSE BUTTONS
        private void CmdExportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                if (mobjSelectedReport == null)
                {
                    MessageBox.Show("Please select a report from the list.");
                    return;
                }
                this.Cursor = Cursors.WaitCursor;
                string reportfilename = string.Empty;
                reportfilename = mobjSelectedReport.Index switch
                {
                    32 or 33 or 35 or 39 or 61 or 62 => ProcessRequest(DisjointedData: true),
                    _ => ProcessRequest(DisjointedData: false),
                };
                if (reportfilename == null || reportfilename == "")
                {
                    MessageBox.Show("The report was not generated.", "Report Not Generated", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    if (string.IsNullOrEmpty(reportfilename)) return;
                    MessageBox.Show($"The report has been generated and exported to Excel.\r\n\r\n{reportfilename}", "Report Generated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    System.Diagnostics.Process.Start("explorer.exe", $"/select,\"{reportfilename}\"");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        string ProcessRequest(bool DisjointedData)
        {
            try
            {
                if (mobjSelectedReport == null) throw new Exception("No report selected.");
                TFReportsSQLNext.Classes.SQLSelector sqlSelector = new(mobjSelectedReport, mobjReports);
                if (DisjointedData)
                {
                    sqlSelector.PrepareDisjointedData();
                }
                else
                {
                    ds = sqlSelector.PrepareDataSet();
                    if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0) throw new Exception("No data found for the selected report and options.");
                }

                string reportfilename = SelectFileName(mobjReports.Filename(mobjSelectedReport.Index));// sqlSelector.ReportFileName);
                TFSpreadSheetsNext.TFRCommon tfrCommon = new(mobjSelectedReport, mobjReports, ds, sqlSelector.ReportTitle, reportfilename);
                tfrCommon.ExportToExcel();
                return reportfilename;
            }
            catch (Exception)
            {

                throw;
            }

        }
        string SelectFileName(string reportfilename)
        {
            SaveFileDialog saveFileDialog = new()
            {

                Filter = "Excel Files (*.xlsx)|*.xlsx",
                DefaultExt = "xlsx",
                Title = "Save Report As",
                FileName = reportfilename
            };
            if (saveFileDialog.ShowDialog() == DialogResult.Cancel)
            {
                throw new Exception("Export cancelled by user.");
            }
            return saveFileDialog.FileName;
        }
        #endregion
        #region TREEVIEW
        private void TrvReportSelector_AfterSelect(object sender, TreeViewEventArgs e)
        {
            try
            {
                if (e.Node == null) return;
                if (!mflgLoading)
                {
                    if (e.Node.Parent == null)
                    {
                        mobjSelectedReport = null;
                    }
                    else
                    {
                        mobjSelectedReport = mobjReports[int.Parse(e.Node.Name)];
                    }
                    bool pFlagLoading = mflgLoading;
                    mflgLoading = true;
                    mobjReports.ClearOptions();
                    Frames_PrepareFrames();
                    PrepareOptions();
                    mflgLoading = pFlagLoading;
                    ValidateOptions();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        #region CUSTOMER OPTIONS
        private void OptAllCustomers_CheckedChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                mobjReports.SetAllClients();
                ValidateOptions();
            }
        }
        private void CmbCustomers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (cmbCustomers.SelectedItem == null) return;
                TFReportsSQLNext.Classes.Client selectedCustomer = (TFReportsSQLNext.Classes.Client)cmbCustomers.SelectedItem;

                if (selectedCustomer != null)
                {
                    mobjReports.SelectedCustomer = ((TFReportsSQLNext.Classes.Client)cmbCustomers.SelectedItem).ClientCode;
                    mflgLoading = true;
                    optCustomer.Checked = true;
                    mflgLoading = false;
                    ValidateOptions();

                }
            }
        }
        private void CmbCustomerGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (cmbCustomerGroup.SelectedItem == null) return;
                TFReportsSQLNext.Classes.ClientGroup selectedGroup = (TFReportsSQLNext.Classes.ClientGroup)cmbCustomerGroup.SelectedItem;
                mobjReports.SetCustomerGroup(selectedGroup.Id, selectedGroup.GroupName);
                mflgLoading = true;
                optCustomerGroup.Checked = true;
                mflgLoading = false;
                ValidateOptions();
            }
        }
        #endregion
        #region BSP OPTIONS
        private void OptInternational_CheckedChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                mobjReports.MonthDomestic = false;
                ValidateOptions();
            }
        }
        private void OptDomestic_CheckedChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                mobjReports.MonthDomestic = true;
                ValidateOptions();
            }
        }
        //private void LstBSPMonth_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (!mflgLoading)
        //    {
        //        if (lstBSPMonth.SelectedItem != null)
        //        {
        //            mobjReports.BSPMonthDate = lstBSPMonth.SelectedItem.ToString();
        //        }
        //        else
        //        {
        //            mobjReports.BSPMonthDate = null;
        //        }
        //        ValidateOptions();
        //    }
        //}
        //private void LstBSPFortnight_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    if (!mflgLoading)
        //    {
        //        if (lstBSPFortnight.SelectedItem != null)
        //        {
        //            mobjReports.BSPFortDate = (DateTime)lstBSPFortnight.SelectedItem;
        //        }

        //        ValidateOptions();
        //    }
        //}
        #endregion
        #region DATE OPTIONS
        private void DtpDate1From_ValueChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                mobjReports.Date1From = dtpDate1From.Value;
                mobjReports.Date1Checked = dtpDate1From.Checked;
                ValidateOptions();
            }
        }
        private void DtpDate1To_ValueChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                mobjReports.Date1To = dtpDate1To.Value;
                if (mobjSelectedReport != null && mobjSelectedReport.Date2Init == ReportsNext.ReportsItem.DateInitValue.From5DaysAfterEOM)
                {
                    SetDate2FromFor5DaysAfterEOM();
                }
                ValidateOptions();
            }
        }
        private void SetDate2FromFor5DaysAfterEOM()
        {
            mobjReports.Date1Checked = dtpDate1To.Checked;
            mobjReports.Date2From = new DateTime(mobjReports.Date1To.Year, mobjReports.Date1To.Month + 1, 0).AddDays(5);
            mflgLoading = true;
            dtpDate2From.Value = mobjReports.Date2From;
            mflgLoading = false;
        }
        private void DdtpDate2From_ValueChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                mobjReports.Date2From = dtpDate2From.Value;
                mobjReports.Date2Checked = dtpDate2From.Checked;
                ValidateOptions();
            }
        }
        private void DtpDate2To_ValueChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                mobjReports.Date2To = dtpDate2To.Value;
                ValidateOptions();
            }
        }
        private void LstYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (lstYear.SelectedItem != null)
                {
                    mobjReports.ReportYear = (int)lstYear.SelectedItem;
                }
                ValidateOptions();
            }
        }
        private void LstMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (lstMonth.SelectedItem != null)
                {
                    mobjReports.ReportMonth = (int)lstMonth.SelectedItem;
                }
                ValidateOptions();
            }
        }
        #endregion
        #region CHECKBOX AND RADIO BUTTON OPTIONS
        private void ChkCheckOption1_CheckedChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (mobjSelectedReport != null)
                {
                    mobjReports.BooleanOption1 = chkCheckOption1.Checked;
                    ValidateOptions();
                }
            }
        }
        private void OptTriplet0_CheckedChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (mobjSelectedReport != null)
                {
                    mobjReports.OptionTriplet = 0;
                    ValidateOptions();
                }
            }
        }
        private void OptTriplet1_CheckedChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (mobjSelectedReport != null)
                {
                    mobjReports.OptionTriplet = 1;
                    ValidateOptions();
                }
            }
        }
        private void OptTriplet2_CheckedChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (mobjSelectedReport != null)
                {
                    mobjReports.OptionTriplet = 2;
                    ValidateOptions();
                }
            }
        }

        #endregion
        #region TEXT ENTRY
        private void TxtTextEntry_TextChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (mobjSelectedReport != null)
                {
                    mobjReports.TextEntry = txtTextEntry.Text;
                    ValidateOptions();
                }
            }
        }
        #endregion
        #region LSTGROUP
        private void LstGroupList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!mflgLoading)
            {
                if (lstGroupList.SelectedItem != null)
                {
                    mobjReports.GroupList = lstGroupList.Text;
                }
                ValidateOptions();
            }
        }
        #endregion
    }
}