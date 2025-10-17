using System;
namespace TFReportsNext
{
    public partial class FormTFReportsNext 
    {
        const string VersionText = "V 2025.16 17/09/2025";

/*
V 2025.15 17/09/2025               :  added report 67 Columbia for Bellina
V 2025.14 05/09/2025               :  Report 38 - added 2 new columns Supplier Code and Supplier Name
V 2025.13 07/07/2025               :  Report 66 - There was a problem calculating the Indeces
V 2025.12 17/06/2025               :  Changed report 63 to use Invoice Issue Date instead of Commercial Transaction Date
V 2025.11 22/05/2025               :  Changed Report 66 to use from/to date rather than specific month
V 2025.10 22/05/2025               :  Added Report 66
V 2025.09 15/05/2025               :  Added Report 64
V 2025.08 02/05/2025               :  Report 63 Temenos
V 2025.07 28/03/2025               :  Added invoice currency option to report 22
V 2025.06 19/03/2025               :  added report 61 and 62
V 2025.05 21/02/2025               :  fixed 28 tabs with group details were not sorting properly
V 2025.04 03/02/2025               :  added booking date and other minor details to 60
V 2025.03 30/01/2025               :  added report 60
V 2025.02 30/01/2025               :  added With Fare Basis option to 55
V 2025.01 24/01/2025               :  added By Client sheet to 28
V 2024.20 10/12/2024               :  added Inactive to Report 12
V 2024.19 04/12/2024               :  Reports.E28_ActionReasonItem -> public int Ranking(int index)
                                   :  added division by zero check
V 2024.18 18/11/2024               :  fixed Potential downsell in report 28. It was showing pax total and not PNR total
V 2024.17 22/10/2024               :  added column Net Remit to report 30
V 2024.16 11/09/2024               :  added report 57 TUI 030366
V 2024.15 08/08/2024               :  added Report parameter which sets initial values for the report
V 2024.14 08/08/2024               :  add TOTAL row for report 28 sheet Action Totals
V 2024.13 06/08/2024               :  added analysis per hour to report 28 Action Totals sheet
V 2024.12 17/07/2024               :  created all the worksheets in report 28 - Optimization Savings
V 2024.11 27/06/2024               :  added a check for string length
                                   :  switch ($"{CostCentre}0000".Substring(0, 4))
                                   :  in E35_Item.TransactionKey
V 2024.10 27/06/2024               :  added Count of Passive and non-air to report 32
V 2024.09 26/06/2024               :  Changed report 28 to show POSTPONE/REJECT reason
V 2024.08 20/06/2024               :  added report 56 - Clients per Group
V 2024.07 12/06/2024               :  add option ignore OMIT/VOID to report 18
                                   :  I want this for the lowest class of service report in TicketStatusNext
V 2024.06 08/05/2024               :  fixed optimization reports 49 & 50 to read from correct SQL server
V 2024.05 04/04/2024               :  added report 54
V 2024.04 29/03/2024               :  added report 53
V 2024.03 29/03/2024               :  fixed report 32
                                   :  if there was no reason for travel the program crashed
V 2024.02 20/02/2024               :  added report 51
V 2024.01 05/01/2024               :  added reports 49 & 50
V 2023.28 20/12/2023               :  added omit option to report 47
                                   :  added report 48
V 2023.27 13/12/2023               :  added details to the IW5 columns in report 47
V 2023.26 12/12/2023               :  added IW5 columns to report 47
V 2023.25 08/12/2023               :  added report 47 Profit Report Totals
                                   :  added Ops group selector to report 12
V 2023.24 20/11/2023               :  added cabin class and duration to report 45
V 2023.23 20/11/2023               :  Changed Stored Procedures for Report 43 to 43a and 19 to 19b
V 2023.22 06/11/2023               :  added report 19a
                                   :  added Client Group and Ops Team to report 19
V 2023.21 09/10/2023               :  added fare basis to report 45
V 2023.20 15/09/2023               :  added report 45
V 2023.18 05/09/2023               :  added report 43 Daily Profit with Provisional Analysis
V 2023.17 18/08/2023               :  Fixed bug where date2 checkbox was reset when chkCheck was clicked
V 2023.16 20/07/2023               :  Added new option to Report 23 - Separate Worksheet per Booked By
V 2023.15 12/07/2023               :  Added report 42 Air Ticket with FC
                                   :  Added FC Column to Report 19
V 2023.14 12/06/2023               :  added report 41 Profit Report with RINVA Analysis
V 2023.13 01/06/2023               :  Added new columns to Report 32 (Dummy/FixedMarkup/Void/Refund)
V 2023.12 25/05/2023               :  Fix RXRegex in PNRHistory.Remarks
V 2023.11 17/05/2023               :  add new columns to 18 (sell, buy, profit, pax+-)
V 2023.10 12/05/2023               :  Revert Report 23 to old version
                                   :  Add Report 40
V 2023.9 08/05/2023                :  Revert report 12 to previous version
V 2023.8 27/4/2023                 :  Finish off converting to Stored Procedures and adding to log file
V 2023.7 27/4/2023                 :  Added report 39 and added all client profile elements to report 30
V 2023.6 21/3/2023                 :  converting all SQL to stored procedures and adding log file
V 2023.5 10/3/2023                 :  missed a column in report 38. Now fixed
V 2023.4 9/3/2023                  :  added stored procedure for 19 and 37
V 2023.3 9/3/2023                  :  added reports 37 and 38
V 2023.2 6/3/2023                  :  fixed NextDB connection for report 32
V 2022.12 16/2/2023                :  added client number and some other fields to report 36
V 2022.11 15/2/20230               :  added report 36 which is a clone of 23 and works for all Sea Chefs Business Units automatically
V 2022.10 15/2/20230               :  
V 2022.09 14/2/2023                :  added reports 33 & 35 for Gaslog
V 2022.08 30/1/2023                :  Changed report 29 & 31 to select supplier name depending on client code
                                   :  created stored procedures:
                                   :  ATPIData.dbo.TFReports_E29_SeaChefsDetailed
                                   :  ATPIData.dbo.TFReports_E31_SeaChefsStatementCheck
V 2022.07 27/1/2023                :  added Client Team column to report 18
V 2022.06 27/1/2023                :  added Other Services column to report 18
V 2022.05 29/12/2022               :  Added Report 32 Accounting QC for Selling Price
                                   :  added new projects PNRHistory and Reports
V 2022.04 08/10/2022               :  Removed DocType 134 (RINVA) from the excluded document types 15 Profit Report
V 2022.03 02/03/2022               :  Exclude Doc Type 109 in Report 15. Alos exclude clients from AmadeusReports.dbo.TFReportExclude
V 2022.02 19/01/2022               :  Added report 31
V 2022.01 10/01/2022               :  Changed report 23 to check for last departure date and not first departure date
V 2021.25 27/12/2021               :  Problem with null values in report 19
V 2021.24 24/11/2021               :  added cancelled doc analysis to report 18
V 2021.23 20/10/2021               :  SQL 30 added
V 2021.22 29/09/2021               :  SQL 23 - added a new table for Sea Chefs Business Units
V 2021.21 22/09/2021               :  Added one more WHERE clause in SQL 23 to avoid accounting entries AND DocTypes.UsesFSD = 1
V 2021.20 21/09/2021               :  added more codes to SQL for 23
V 2021.19 13/09/2021               :  Fixed 23 to not have any NULL values so that it can run for individual Client Codes
V 2021.17 18/08/2021               :  Customer Group SQL Command had an infinite loop because of the brackets in "With Customer_Group()"
                                   :  I removed all With statements in SQLCommands.vb
V 2021.16 29/07/2021               :  Supplier number also changed in SQL Query for 23
V 2021.15 29/07/2021               :  forgot two END statement in sql QUERY
V 2021.14 29/07/2021               :  Added new BUs to 23 Sea Chefs
V 2021.13 28/07/2021               :  Changed 23 to cater for a different structure of the client groups and Sea Chefs Business Units
V 2021.12 23/07/2021               :  change 23 for the layout changes requested by Sea Chefs
                                   :  also introduced a parameter to limit the selectable Customer Groups when running the Sea Chefs report
V 2021.11 12/07/2021               :  fixed SQL 15 and 19 to prevent duplicate cost of hotels in some cases
V 2021.10 08/07/2021               :  fixed SQL 23 to avoid NULL references
V 2021.9 01/07/2021                :  fixed SQL 29 to avoid NULL references
V 2021.8 22/06/2021                :  renamed report 23 to 29 Sea Chefs Detailed
                                   :  and redesigned report 23 to match Sea Chefs latest requirements
V 2021.7 12/05/2021                :  added report 28 - Optimization Savings
V 2021.6 11/05/2021                :  changed the SQL query for 15 Daily Profit Report to separate air tickets from services
V 2021.5 07/05/2021                :  Changed the calculation of Provisional Profit from uninvoiced Pax in report 15
                                   :  added extra checks for cancellation documents in report 23
V 2021.4 23/03/2021                :  Added departure date to report 22 - Euronav
V 2021.3 11/03/2021                :  Changed the query for report 23 to check only for last segment departure date
V 2021.1 09/03/2021                :  Added reports 24 and 25 for Profit per agent
V 2020.9 175/02/2021               :  23 Seachefs Statement. Changed the Description Column to add Booked By and TRID
                                   :  Added 2 columns, client code and client name
V 2020.8 15/02/2021                :  Changed SQL for 23 to cater for multiple records returned by subquery
                                   :  also added timeout=120 to avoid timeout problem
V 2020.7 10/02/2021                :  Changed the form. Removed tool strip button and made it normal button because tool strip does not cause validation and confuses the user.
V 2020.6 10/02/2021                :  Fixed the SQL statement for report 23. It was not checking the client number correctly.
V 2020.5 02/02/2021                :  Small error in Filename for report 23 fixed
V 2020.4 02/02/2021                :  23. Sea Chefs statement
V 2020.3 04/09/2020                :  15. Daily Report. Client groups had incorrect YTD Pax because it ignored client codes that were active during the year but inactive in the requested period
                                   :  added "OR #TempTableYTD.PaxAIR<>0 OR #TempTableYTD.PaxServices <> 0" in the SQL where client codes are selected
V 2020.2 23/07/2020                :  Some SQL queries were still referring to the server EUDC-SQLCL14. This has now been fixed
V 2020.1 22/07/2020                :  Added new column "Provisional Uninvoiced T/O" in report 15-Daily Profit
V 2020.0 30/06/2020                :  The new generation
                                   :  using port 453334 to connect to the SQL database
01/06/2020 (TReports_2_0_0_61)     :  Report 22-Euronav - Added WHERE to exclude cancellation documents and omitted documents
29/05/2020 (TReports_2_0_0_60)     :  Report 22-Euronav - Change the SQL statement to ignore cancelled documents
29/05/2020 (TReports_2_0_0_59)     :  Report 22-Euronav - Change the SQL statement to check invoice date and not commercial transaction date
22/05/2020 (TReports_2_0_0_58)     :  Changed the DB Connection for the new Panasoft database
25/02/2020 (TFReports_2_0_0_57)    :  Added 2 new columns to report 18. Connected Document and Passenger Remarks
                                   :  Created Report 22-Euronav which is similar to 00-Euronav but is based on report 18
17/02/2020 (TFReports_2_0_0_56)    :  Add column Account Code in report 18
10/02/2020 (TFReports_2_0_0_55)    :  Remove the check CommercialTransactions.CustomDescription3 <> from reports 15 and 19
21/11/2019 (TFReports_2_0_0_54)    :  Added fields "Departure Date" and "Arrival Date" to report 18
29/10/2019 (TFReports_2_0_0_53)    :  Added field "Reference" to report 18
12/07/2019 (TFReports_2_0_0_50)    :  Added Issuing Airline to Report 19
                                   :  Added New Group "Optimize Reports" and new report "21 Report by Verified User"
                                   :  This required the addition of new parameters
                                   :  1 - to access the AmadeusReports database and
                                   :  2 - to reuse the listbox in the options to show the Agents
24/06/2019 (TFReports_2_0_0_47)    :  Report 18 - All Customers option was not working and I also aded one more column with the routing
                                   :  Report 18 - Added a new option to select airlines
                                   :  Changed text box for Airline codes to CAPITALS
                                   :  after a file is created, the directory is opened
24/06/2019 (TFReports_2_0_0_46)    :  Added 3 new columns in report 18 - Salesman, Issuing Agent, Creator Agent
21/06/2019 (TFReports_2_0_0_45)    :  Added a new column in report 18 - Ticketing Airline
20/06/2019 (TFReports_2_0_0_44)    :  SQL Statement for report 18 crashed when selecting client group because SelectedCustomer is nothing
12/06/2019 (TFReports_2_0_0_43)    :  Added a new option in the 2 From/To dates. The checkbox allows the user to ignore one of the two dates. If he ignores both, the options are not valid and the program will not run.
                                   :  Specifically in the case of report 20. Hellas Confidence, the user can select From/To Issue Date and/or From/To Invoice Date.
                                   :  The SQL for report 20 Hellas Confidence was changed to take into account these 2 options
29/05/2019 (TFReports_2_0_0_42)    :  1. Report 18 was highlighting the omitted invoices incorrectly
                                   :  Spreadsheet.E19_DailyProfitReportInvoicesWithTicketNumber()
                                   :  the line : If mdsDataSet.Tables(0).Rows(i).Item(60) Then
                                   :  was changed to : If mdsDataSet.Tables(0).Rows(i).Item(63) Then
27/05/2019 (TFReports_2_0_0_41)    :  1. added report 20 Hellas Confidence. This is a cut-down version of report 18.
03/05/2019 (TFReports_2_0_0_40)    :  1. Added Uninvoiced Pax and Provisional Profit in 15 Daily Profit Report
24/04/2019 (TFReports_2_0_0_39)    :  1. Corrected an error in report 19 where the IW amount was not being multiplied by the currency rate
                                   :  2. Added new report 17 Service Fee Analysis
22/04/2019 (TFReports_2_0_0_38)    :  1. Added IW11 to reports 15 and 19
                                   :  2. Added a new option to report 19 "With tickets" and removed report 17 which is the same as old report 19 but without tickets
12/04/2019 (TFReports_2_0_0_37)    :  1. Added a new option "All clients"
10/04/2019 (TFReports_2_0_0_36)    :  1. Added report 19 which is a copy of report 17 with extra columns for Ticket Number, Face Value, Taxes, Commission, Discount, Cancellation Fee and TF
06/03/2019 (TFReports_2_0_0_34)    :  1. Report 15 Daily Profit Report
                                   :  Increase the timeout to 240
13/02/2019 (TFReports_2_0_0_33)    :  1. Report 12 Profit Per Client Group with Budget Comparison
                                   :  Add a new category 99 - OTHER to show budget for entries not connected to actual clients (new clients, expansion for existing clients, etc
                                   :  This required the addition of 2 new fields in the TFClientBudget table for the Code and Name that will be shown in the report
11/02/2019 (TFReports_2_0_0_32)    :  1. Report 18 Air Ticket Sales
                                   :  Non Commercial Transactions RINVA, Etc do not have entry in CommercialTransactionCards and give NULL values that were not trapped in thq SQL query causing the program to crash
*/
    }
}