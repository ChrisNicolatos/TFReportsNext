using Microsoft.Data.SqlClient;
using System;
using System.Data;
namespace TFReportsNext
{
    public partial class FormTFReportsNext
    {
        private void Frames_PrepareFrames()
        {

            /*
            ' Options from Top to Bottom are:
            ' Customer Frame
            ' Options Frame
            ' in the Options frame we have
            ' Date1 From/To (with 3 labels for date description/from/to)
            ' Date2 From/To (with 3 labels for date description/from/to)
            ' CheckOption1
            ' opt02International
            ' opt02Domestic
            ' and then the list boxes/text boxes
            ' These are from left to right (and each one with its appropriate heading label):
            ' lstYear/lstMonth
            ' txtTicketNumbers
            ' lstBSPMonth
            ' lstBSPFortnight
            ' lstAirlines
            */
            try
            {
                if (mobjSelectedReport != null)
                {
                    pLocX = 16;
                    pLocY = 25;
                    pStepX = 25;
                    pStepY = 25;
                    fraReportOptions.Visible = false;
                    fraReportOptions.Text = mobjSelectedReport.ReportName;

                    Frames_PrepareCustomer();
                    Frames_PrepareDates();
                    Frames_PrepareIntDom();
                    Frames_PrepareOptionsTriplet();
                    Frames_PrepareCheckBox();
                    Frames_PrepareReportYearMonth();
                    Frames_PrepareTextEntry();
                    //Frames_PrepareBSPMonth();
                    //Frames_PrepareBSPFortnight();
                    Frames_PrepareGroupList();
                    fraReportOptions.Visible = true;
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        private void Frames_PrepareCustomer()
        {
            // Set client selection depending on the reports client code flag
            // None, Client Code and Group, Client Code and Seachefs Group, Client Code Only
            // None =0 : No client selection eneabled
            //ClientCodeOnly =3 : Only client code selection enabled
            //ClientCodeAndGroup =1 : Client code and group selection enabled
            //ClientCodeAndSeachefsGroup =2 : Client code and  group selection enabled
            // (Seachefs group is just a predefined group of customers) when the group is populated
            // only the Seachefs group is shown
            if (mobjSelectedReport == null) return;

            if (mobjSelectedReport.ClientCode == ReportsNext.ReportsItem.ClientCodeSelect.ClientCodeOnly)
            {
                optAllCustomers.Visible = false;
                optCustomerGroup.Visible = false;
                cmbCustomerGroup.Visible = false;
                Frames_PopulateCustomers();
                fraCustomer.Visible = true;
                optCustomer.Checked = true;
            }
            else if (mobjSelectedReport.ClientCode > 0)
            {
                optAllCustomers.Visible = true;
                optCustomerGroup.Visible = true;
                cmbCustomerGroup.Visible = true;
                Frames_PopulateCustomers();
                Frames_PopulateCustomerGroups(mobjSelectedReport.ClientCode);
                fraCustomer.Visible = true;
                optAllCustomers.Checked = true;
            }
            else
            {
                fraCustomer.Visible = false;
            }
        }
        private void Frames_PrepareDates()
        {
            // Dates
            // if there is text in Date1Text then show the date1 controls
            if (mobjSelectedReport == null) return;
            if (mobjSelectedReport.Date1Text.Length == 0)
            {
                lblDate1.Visible = false;
                lblDate1To.Visible = false;
                dtpDate1From.Visible = false;
                dtpDate1To.Visible = false;
            }
            else
            {
                lblDate1.Text = mobjSelectedReport.Date1Text;
                lblDate1.Visible = true;
                lblDate1.Location = new Point(pLocX, pLocY);
                lblDate1To.Visible = true;
                lblDate1To.Location = new Point(lblDate1To.Location.X, pLocY);
                dtpDate1From.ShowCheckBox = mobjSelectedReport.Date1Optional;
                dtpDate1From.Checked = true;
                dtpDate1From.Visible = true;
                dtpDate1From.Location = new Point(dtpDate1From.Location.X, pLocY);
                dtpDate1To.Visible = true;
                dtpDate1To.Location = new Point(dtpDate1To.Location.X, pLocY);
                // the Date1Init flag determines how the date is initialised
                // FirstPrevMonthToEnd = 1, FirstCurrMonthToToday = 2, FirstJanToToday = 3, FromToPrevDayOrFriday = 4
                // FirstCurrMonthToYesterday = 5, From5DaysAfterEOM=6
                // Default is From/To Previous month 1st to last day of month
                // FromToPrevDayOrFriday = 4: If today is Monday then set to previous Friday otherwise set to yesterday                    
                if (mobjSelectedReport.Date1Init == ReportsNext.ReportsItem.DateInitValue.FromToPrevDayOrFriday)
                {
                    if (DateTime.Today.DayOfWeek == DayOfWeek.Monday)
                    {
                        mobjReports.Date1From = DateTime.Today.AddDays(-3);
                    }
                    else
                    {
                        mobjReports.Date1From = DateTime.Today.AddDays(-1);
                    }
                    mobjReports.Date1To = DateTime.Today.AddDays(-1);
                }
                // FirstcurrMonthToToday = 2: From first day of current month to today
                else if (mobjSelectedReport.Date1Init == ReportsNext.ReportsItem.DateInitValue.FirstCurrMonthToToday)
                {
                    mobjReports.Date1Checked = true;
                    mobjReports.Date1From = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                    mobjReports.Date1To = DateTime.Today;
                    if (mobjReports.Date1To < mobjReports.Date1From)
                    {
                        mobjReports.Date1To = mobjReports.Date1From;
                    }
                }
                // firstcurrmonthToYesterday = 5: From first day of current month to yesterday
                else if (mobjSelectedReport.Date1Init == ReportsNext.ReportsItem.DateInitValue.FirstCurrMonthToYesterday)
                {
                    mobjReports.Date1Checked = true;
                    mobjReports.Date1From = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                    mobjReports.Date1To = DateTime.Today.AddDays(-1);
                    if (mobjReports.Date1To < mobjReports.Date1From)
                    {
                        mobjReports.Date1To = mobjReports.Date1From;
                    }
                }
                else // From/To Previous month 1st to last day of month
                {
                    mobjReports.Date1From = (new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1)).AddMonths(-1);
                    mobjReports.Date1To = mobjReports.Date1From.AddMonths(1).AddDays(-1);
                }
                dtpDate1From.Value = mobjReports.Date1From;
                dtpDate1To.Value = mobjReports.Date1To;
                pLocY += pStepY;

            }
            if (mobjSelectedReport.Date2Text == "")
            {
                lblDate2.Visible = false;
                lblDate2To.Visible = false;
                dtpDate2From.Visible = false;
                dtpDate2To.Visible = false;
            }
            else
            {
                lblDate2.Text = mobjSelectedReport.Date2Text;
                lblDate2.Visible = true;
                lblDate2.Location = new Point(pLocX, pLocY);
                dtpDate2From.ShowCheckBox = mobjSelectedReport.Date2Optional;
                dtpDate2From.Checked = true;
                dtpDate2From.Visible = true;
                dtpDate2From.Location = new Point(dtpDate2From.Location.X, pLocY);
                if (mobjSelectedReport.Date2OnlyFrom)
                {
                    lblDate2To.Visible = false;
                    dtpDate2To.Visible = false;
                }
                else
                {
                    lblDate2To.Visible = true;
                    lblDate2To.Location = new Point(lblDate2To.Location.X, pLocY);
                    dtpDate2To.Visible = true;
                    dtpDate2To.Location = new Point(dtpDate2To.Location.X, pLocY);
                }
                if (mobjSelectedReport.Date2Init == ReportsNext.ReportsItem.DateInitValue.From5DaysAfterEOM)
                {
                    mobjReports.Date2From = new DateTime(mobjReports.Date1To.Year, mobjReports.Date1To.Month, 1).AddMonths(1).AddDays(5);
                    mobjReports.Date2To = mobjReports.Date2From;
                }
                dtpDate2From.Value = mobjReports.Date2From;
                dtpDate2To.Value = mobjReports.Date2To;
                pLocY += pStepY;
            }
        }
        private void Frames_PrepareIntDom()
        {
            // Domestic/International options
            if (mobjSelectedReport == null) return;
            opt02Domestic.Visible = mobjSelectedReport.DomInt;
            opt02International.Visible = mobjSelectedReport.DomInt;
            if (mobjSelectedReport.DomInt)
            {
                opt02Domestic.Location = new Point(pLocX, pLocY);
                pLocY += pStepY;
                opt02International.Location = new Point(pLocX, pLocY);
                pLocY += pStepY;
            }

        }
        private void Frames_PrepareOptionsTriplet()
        {
            if (mobjSelectedReport == null) return;
            // Triplet options
            fraOptionsTriplet.Visible = mobjSelectedReport.OptionsTriplet;
            if (mobjSelectedReport.OptionsTriplet)
            {
                fraOptionsTriplet.Location = new Point(pLocX, pLocY);
                optTriplet0.Text = mobjSelectedReport.Options0Text;
                optTriplet1.Text = mobjSelectedReport.Options1Text;
                optTriplet2.Text = mobjSelectedReport.Options2Text;
                pLocY += fraOptionsTriplet.Height + pStepY;
            }
        }
        private void Frames_PrepareCheckBox()
        {
            // Check box options
            if (mobjSelectedReport == null) return;
            if (mobjSelectedReport.CheckBoxText == "")
            {
                chkCheckOption1.Visible = false;
            }
            else
            {
                chkCheckOption1.Visible = true;
                chkCheckOption1.Text = mobjSelectedReport.CheckBoxText;
                chkCheckOption1.Location = new Point(pLocX, pLocY);
                pLocY += pStepY;
            }
        }
        private void Frames_PrepareReportYearMonth()
        {
            if (mobjSelectedReport == null) return;
            if (mobjSelectedReport.ReportYearMonth)
            {
                Frames_LoadYears();
                lblYear.Visible = true;
                lblYear.Location = new Point(pLocX, pLocY);
                lstYear.Visible = true;
                lstYear.Location = new Point(pLocX, pLocY + lblYear.Size.Height);
                lstYear.Height = fraReportOptions.Height - lstYear.Top - pStepY;
                lstYear.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
                pLocX += lstYear.Size.Width;

                lblMonth.Visible = true;
                lblMonth.Location = new Point(pLocX, pLocY);
                lstMonth.Visible = true;
                lstMonth.Location = new Point(pLocX, pLocY + lblMonth.Size.Height);
                lstMonth.Height = fraReportOptions.Height - lstMonth.Top - pStepY;
                lstMonth.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
                pLocX += lstMonth.Size.Width + pStepX;
            }
            else
            {
                lblYear.Visible = false;
                lblMonth.Visible = false;
                lstYear.Visible = false;
                lstMonth.Visible = false;
            }
        }
        private void Frames_PrepareTextEntry()
        {
            if (mobjSelectedReport == null) return;
            if (!string.IsNullOrEmpty(mobjSelectedReport.TextEntry))
            {
                lblTextEntry.Text = mobjSelectedReport.TextEntry;
                lblTextEntry.Visible = true;
                lblTextEntry.Location = new Point(pLocX, pLocY);
                txtTextEntry.Visible = true;
                txtTextEntry.Multiline = mobjSelectedReport.TextEntryMultiLine;
                txtTextEntry.Location = new Point(pLocX, pLocY + lblTextEntry.Size.Height);
                txtTextEntry.Height = fraReportOptions.Height - txtTextEntry.Top - pStepY;
                txtTextEntry.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
                pLocX += txtTextEntry.Size.Width + pStepX;
            }
            else
            {
                lblTextEntry.Visible = false;
                txtTextEntry.Visible = false;
            }
        }
        //private void Frames_PrepareBSPMonth()
        //{
        //    if (mobjSelectedReport == null) return;
        //    if (mobjSelectedReport.BSPMonth)
        //    {
        //        Frames_LoadBSPMonths();
        //        lblBSPMonth.Visible = true;
        //        lblBSPMonth.Location = new Point(pLocX, pLocY);
        //        lblBSPMonth.Text = "BSP Month";
        //        lstBSPMonth.Visible = true;
        //        lstBSPMonth.Location = new Point(pLocX, pLocY + lblBSPMonth.Size.Height);
        //        lstBSPMonth.Height = fraReportOptions.Height - lstBSPMonth.Top - pStepY;
        //        lstBSPMonth.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
        //        pLocX += lstBSPMonth.Size.Width + pStepX;
        //    }
        //    else if (mobjSelectedReport.OptimizationMonth)
        //    {
        //        Frames_LoadVerificationYears();
        //        lblBSPMonth.Visible = true;
        //        lblBSPMonth.Location = new Point(pLocX, pLocY);
        //        lblBSPMonth.Text = "Verified Year";
        //        lstBSPMonth.Visible = true;
        //        lstBSPMonth.Location = new Point(pLocX, pLocY + lblBSPMonth.Size.Height);
        //        lstBSPMonth.Height = fraReportOptions.Height - lstBSPMonth.Top - pStepY;
        //        lstBSPMonth.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
        //        pLocX += lstBSPMonth.Size.Width + pStepX;
        //    }
        //    else
        //    {
        //        lblBSPMonth.Visible = false;
        //        lstBSPMonth.Visible = false;
        //    }
        //}
        //private void Frames_PrepareBSPFortnight()
        //{
        //    if (mobjSelectedReport == null) return;
        //    if (mobjSelectedReport.BSPFortnight)
        //    {
        //        Frames_LoadBSPFortnights();
        //        lblBSPFortnight.Visible = true;
        //        lblBSPFortnight.Location = new Point(pLocX, pLocY);
        //        lstBSPFortnight.Visible = true;
        //        lstBSPFortnight.Location = new Point(pLocX, pLocY + lblBSPFortnight.Size.Height);
        //        lstBSPFortnight.Height = fraReportOptions.Height - lstBSPFortnight.Top - pStepY;
        //        lstBSPFortnight.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
        //        pLocX += lstBSPFortnight.Size.Width + pStepX;
        //    }
        //    else
        //    {
        //        lblBSPFortnight.Visible = false;
        //        lstBSPFortnight.Visible = false;
        //    }
        //}
        private void Frames_PrepareGroupList()
        {
            if (mobjSelectedReport == null) return;
            if (mobjSelectedReport.GroupList != ReportsNext.ReportsCollection.GroupListType.Undefined)
            {
                if (mobjSelectedReport.GroupList == ReportsNext.ReportsCollection.GroupListType.OperatorsGroup)
                {
                    // loadOperationsGroup();
                }
                else
                {
                    Frames_LoadAgents();
                }
                lblGroupList.Text = mobjSelectedReport.GroupListText;
                lblGroupList.Visible = true;
                lblGroupList.Location = new Point(pLocX, pLocY);
                lstGroupList.Visible = true;
                lstGroupList.Location = new Point(pLocX, pLocY + lblGroupList.Size.Height);
                lstGroupList.Height = fraReportOptions.Height - lstGroupList.Top - pStepY;
                lstGroupList.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom;
                pLocX += lstGroupList.Size.Width + pStepX;
            }
            else
            {
                lblGroupList.Visible = false;
                lstGroupList.Visible = false;
            }
        }
        private void Frames_PopulateCustomers()
        {
            try
            {
                if (cmbCustomers.DataSource == null)
                {
                    TFReportsSQLNext.Classes.ClientDataSet objClients = new();
                    SqlDataAdapter da = new(objClients.ClientDataSetCmd());
                    DataSet ds = new();
                    da.Fill(ds);
                    cmbCustomers.DataSource = ds.Tables[0];
                    cmbCustomers.DisplayMember = "DispName";
                    //objClients.Load();
                    //foreach (TFReportsSQLNext.Classes.Client client in objClients)
                    //{
                    //    cmbCustomers.Items.Add(client);
                    //}
                    //if (cmbCustomers.Items.Count > 0)
                    //{
                    //    cmbCustomers.SelectedIndex = 0;
                    //}
                    ValidateOptions();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        //private void Frames_PopulateCustomers()
        //{
        //    try
        //    {
        //        if (cmbCustomers.Items.Count == 0)
        //        {
        //            TFReportsSQLNext.Classes.ClientList objClients = [];
        //            objClients.Load();
        //            foreach (TFReportsSQLNext.Classes.Client client in objClients)
        //            {
        //                cmbCustomers.Items.Add(client);
        //            }
        //            if (cmbCustomers.Items.Count > 0)
        //            {
        //                cmbCustomers.SelectedIndex = 0;
        //            }
        //            ValidateOptions();
        //        }
        //    }
        //    catch (Exception)
        //    {
        //        throw;
        //    }
        //}
        private void Frames_PopulateCustomerGroups(ReportsNext.ReportsItem.ClientCodeSelect pClientCode)
        {

            try
            {
                if (cmbCustomerGroup.Items.Count == 0)
                {
                    TFReportsSQLNext.Classes.ClientGroupList objClients = [];
                    objClients.Load(pClientCode);
                    foreach (TFReportsSQLNext.Classes.ClientGroup clientgroup in objClients)
                    {
                        cmbCustomerGroup.Items.Add(clientgroup);
                    }
                    if (cmbCustomerGroup.Items.Count > 0)
                    {
                        cmbCustomerGroup.SelectedIndex = 0;
                    }
                    ValidateOptions();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private void Frames_LoadYears()
        {
            try
            {
                if (lstYear.Items.Count == 0)
                {
                    for (int i = DateTime.Today.Year; i >= LOADYEARSMINYEAR; i--)
                    {
                        lstYear.Items.Add(i);
                    }
                    lstYear.SelectedIndex = 0;

                    lstMonth.Items.Clear();
                    for (int i = 1; i <= 12; i++)
                    {
                        lstMonth.Items.Add(i);
                    }
                    lstMonth.SelectedIndex = DateTime.Today.Month - 1;
                    mobjReports.ReportYear = Convert.ToInt32(lstYear.SelectedItem);
                    ValidateOptions();
                }

            }
            catch (Exception)
            {

            }
        }
        private void Frames_LoadAgents()
        {
            if (lstGroupList.Items.Count == 0)
            {
                TFReportsSQLNext.Classes.Agents objAgents = [];
                objAgents.Load();
                foreach (string agent in objAgents)
                {
                    lstGroupList.Items.Add(agent);
                }
                if (lstGroupList.Items.Count > 0)
                {
                    lstGroupList.SelectedIndex = 0;
                }
            }
        }
        //private void Frames_LoadBSPMonths()
        //{
        //    try
        //    {
        //        if (lstBSPMonth.Items.Count == 0)
        //        {
        //            for (int i = DateTime.Today.Year; i >= LOADBSPMONTHSMINYEAR; i--)
        //            {
        //                for (int j = 12; j >= 1; j--)
        //                {
        //                    lstBSPMonth.Items.Add(i.ToString() + j.ToString("00"));
        //                }
        //                lstBSPMonth.Items.Add(i);
        //            }
        //            lstBSPMonth.SelectedIndex = 0;
        //        }
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //}
        //private void Frames_LoadBSPFortnights()
        //{
        //    if (lstBSPFortnight.Items.Count == 0)
        //    {
        //        TFReportsSQLNext.Classes.BSPFortnights objBSPFortnights = [];
        //        objBSPFortnights.Load();
        //        foreach (DateTime dt in objBSPFortnights)
        //        {
        //            lstBSPFortnight.Items.Add(dt.ToString("yyyy-MM-dd"));
        //        }
        //        if (lstBSPFortnight.Items.Count > 0)
        //        {
        //            lstBSPFortnight.SelectedIndex = 0;
        //        }
        //    }
        //}
        //private void Frames_LoadVerificationYears()
        //{
        //    try
        //    {
        //        if (lstBSPMonth.Items.Count == 0)
        //        {
        //            for (int i = DateTime.Today.Year; i >= LOADBSPMONTHSMINYEAR; i--)
        //            {
        //                lstBSPMonth.Items.Add(i.ToString());
        //                lstBSPMonth.Items.Add(i);
        //            }
        //            lstBSPMonth.SelectedIndex = 0;
        //        }
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //}
    }
}