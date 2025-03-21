using DocumentFormat.OpenXml.Spreadsheet;
using IronXL;
using System;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using Microsoft.Playwright;
using NHibernate.Mapping.ByCode;
using OpenQA.Selenium.Support.UI;
using OptionCSMSAutomationPlayWright.Model;
using PuppeteerSharp;
using IElementHandle = Microsoft.Playwright.IElementHandle;
using IPage = Microsoft.Playwright.IPage;
using Page = DocumentFormat.OpenXml.Spreadsheet.Page;
using OfficeOpenXml.Style;
using Color = System.Drawing.Color;
using IronXL.Styles;
using ClosedXML.Excel;

namespace OptionCSMSAutomationPlayWright.SISPages
{
    public class MMAuditPages : BasePageObject
    {
        private IPage _page;

        private readonly string directoryPath = "D:\\File";
        public string path = string.Empty;



        public MMAuditPages(IPage page) : base(page)
        {
            _page = page;
            path = Path.Combine(directoryPath, "Audit Summary.txt");

        }

        #region Xpath Elements
        public ILocator TxtAcuUserName => _page.Locator("//*[@id='username']");
        public ILocator TxtAcuPassword => _page.Locator("//*[@id='password']");
        public ILocator BtnAcuLogin => _page.Locator("//*[@id='Login']");
        public ILocator MenuSchool => _page.Locator("(//a[@id='liSchool'])[1]");
        public ILocator TxtAcuSchoolSearch => _page.Locator("//div[@id='example_filter']//following::input[@type='search']");
        public ILocator IconNextGen => _page.Locator("//table[@id='example']//following::a[@title='NextGen']");
        public ILocator TabMMCustom => _page.Locator("//a[text()='Custom']");
        public ILocator TxtAcutisSearch => _page.Locator("//div[@id='DataTables_Table_0_filter']//input[@type='search']");
        public ILocator TxtAcCreditCardAmount => _page.Locator("//*[@id='DataTables_Table_0']/tbody/tr/td[2]");
        public ILocator TxtAcCreditCardCount => _page.Locator("//*[@id='DataTables_Table_0']/tbody/tr/td[3]");
        public ILocator TxtAcCCServiceFeeAmount => _page.Locator("//*[@id='DataTables_Table_0']/tbody/tr/td[4]");
        public ILocator TxtAcCCServiceFeeCount => _page.Locator("//*[@id='DataTables_Table_0']/tbody/tr/td[5]");
        public ILocator TxtAceCheckAmount => _page.Locator("//*[@id='DataTables_Table_0']/tbody/tr/td[6]");
        public ILocator TxtAceCheckCount => _page.Locator("//*[@id='DataTables_Table_0']/tbody/tr/td[7]");
        public ILocator TxtAceCheckServiceFeeAmount => _page.Locator("//*[@id='DataTables_Table_0']/tbody/tr/td[8]");
        public ILocator TxtAceCheckServiceFeeCount => _page.Locator("//*[@id='DataTables_Table_0']/tbody/tr/td[9]");

        // School MM Dashboard Elements
        public ILocator TabMMDashboard => _page.Locator("//div[@id='ExportHtml']//following::a[@class='btn btn-primary' and @href='/fee-dashboard']");
        public ILocator TabCustom => _page.Locator("//div[@id='mmdashboard']//following::a[text()='Custom']");
        public ILocator TxtStartDate => _page.Locator("//input[@id='txtStartDate']");
        public ILocator TxtEndDate => _page.Locator("//input[@id='txtEndDate']");
        public ILocator BtnCustomFilter => _page.Locator("//input[@id='btnpost']");
        public ILocator LblOrgName => _page.Locator("(//table[@id='DataTables_Table_0']//following::tr[@role='row'][2]/td[1])[1]");
        public ILocator LblCCAmount => _page.Locator("(//table[@id='DataTables_Table_0']//following::tr[@role='row'][2]/td[2])[1]");
        public ILocator LblCCCount => _page.Locator("(//table[@id='DataTables_Table_0']//following::tr[@role='row'][2]/td[3])[1]");
        public ILocator LbleCheckAmount => _page.Locator("(//table[@id='DataTables_Table_0']//following::tr[@role='row'][2]/td[4])[1]");
        public ILocator LbleCheckCount => _page.Locator("(//table[@id='DataTables_Table_0']//following::tr[@role='row'][2]/td[5])[1]");
        public ILocator DdlPaymentMethod => _page.Locator("//*[@id='select2-ddlpaymenttype-container']");
        public ILocator SelectCreditCard => _page.Locator("(//li[@class='select2-results__option'])[1]");
        public ILocator SelecteCheck => _page.Locator("(//li[@class='select2-results__option'])[2]");

        // School Ledger Web Elements
        public ILocator TabAction => _page.Locator("(//div[@id='ExportHtml']//following::a[@href='/billing-ledger-view'])[1]");
        public ILocator TabSchoolLedger => _page.Locator("//div[@id='dvExport']//following::a[@href='/billing-ledger-view']");
        public ILocator ChkAlumniDisableduser => _page.Locator("//div[@id='ExportHtml1']//following::div[@class='radioncustom_style form-group']");
        public ILocator DdlFamilyUser => _page.Locator("//div[@id='familydiv']//following::span[@id='select2-UsersFamily-container']");
        public ILocator SelectAll => _page.Locator("//li[text()='(All Families)']");
        public ILocator BtnLedgerFilter => _page.Locator("//button[@id='btnFilter']");
        public ILocator LblTotalDebit => _page.Locator("//*[@id='tblLedger_wrapper']/div[2]/div/div[3]/div/table/tfoot/tr[1]/td[2]");
        public ILocator LblTotalCredit => _page.Locator("//*[@id='tblLedger_wrapper']/div[2]/div/div[3]/div/table/tfoot/tr[1]/td[1]");
        public ILocator LblFamily => _page.Locator("//label[@id='familydiv']/label");
        public ILocator DdlUser => _page.Locator("//*[@id='select2-Users-container']");
        public ILocator LblUser => _page.Locator("//*[@id='userdiv']/label");
        public ILocator SelectAllUsers => _page.Locator("//li[text()='(All)']");
        public ILocator TxtNote => _page.Locator("//*[@id='txtNote']");
        public ILocator TxtReference => _page.Locator("//*[@id='txtReference']");

        // Dashboard Page Elements
        public ILocator AdminDashboard => _page.Locator("//*[@id='sec-slider']/div");
        public ILocator StaffDashboard => _page.Locator("//div[@class='slider-overlay']");
        public ILocator IconUser => _page.Locator("//i[@class='fas fa-user']");
        public ILocator ExpandUserDropdown => _page.Locator("//*[@class='dropdown-menu dropdown-menu-right show' and @x-placement='bottom-end']");
        public ILocator NavigateToParentPortal => _page.Locator("//a[@class='family-link']");
        public ILocator NavigateToStaffPortal => _page.Locator("//a[text()='Staff Portal']");
        public ILocator AdministrationMenu => _page.Locator("//a[@id='li_administration']");
        public ILocator FeeManagementMenu => _page.Locator("//*[@id='li_billing_ledger_view']");
        public ILocator FeeDashboard => _page.Locator("//div[@class='col-md-12 topmenu_action']//a[@href='/fee-dashboard']");
        public ILocator FeeReports => _page.Locator("//div[@class='col-md-12 topmenu_action']//a[@href='/report-list']");
        public ILocator FeeReportsSearch => _page.Locator("//input[@class='form-control input-small input-inline']");
        public ILocator FeeReportsFirst => _page.Locator("//table//tbody//td[@class='sorting_1'][1]//a");
        public ILocator TransStartDateFilter => _page.Locator("//*[@id='txtStartDate']");
        public ILocator TransEndDateFilter => _page.Locator("//*[@id='txtEndDate']");
        public ILocator StaffCheck => _page.Locator("//*[@id='chkStaff']");
        public ILocator BtnFilter => _page.Locator("//input[@onclick='ShowReportDetails(this.id)']");
        public ILocator Btn911Filter => _page.Locator("//input[@id='btnSubmit']");
        public ILocator ReportValue => _page.Locator("//tfoot//td[2]//strong");
        public ILocator Report911Value => _page.Locator("//tfoot//td[2]");
        public ILocator Table => _page.Locator("//*[@id='DataTables_Table_0']");
        public ILocator Table911 => _page.Locator("//*[@id='Table3']");
        public ILocator ReportResult => _page.Locator("//*[@id='export-content']");
        public ILocator RecordLength => _page.Locator("//*[@class='form-control input-xsmall input-inline']");
        public ILocator AllRecord => _page.Locator("//*[@class='form-control input-xsmall input-inline']//option[@value='-1']");
        public ILocator AccountCreatedRadio => _page.Locator("//input[@id='primary1']");
        public ILocator BtnRunReport => _page.Locator("//input[@value='Run Report']");
        public ILocator SchoolName => _page.Locator("//div[@class='oc-school-name']");
        public ILocator LblFundingAmount => _page.Locator("//*[@id='DataTables_Table_0']/tfoot/tr/td[2]/strong");
        public ILocator LblTransactionAmount => _page.Locator("//*[@id='DataTables_Table_0']/tfoot/tr/td[4]/strong");

        // Logout Elements
        public ILocator LinkSignout => _page.Locator("//a[@class='dropdown-item signout']");


        #endregion

        //Methods for passing the user credentials
        //1. Load URL method    
        public async Task LoadURLAsync(string URL)
        {
            await _page.GotoAsync(URL);
            //await _page.SetViewportSizeAsync(1920, 1080);
            Console.WriteLine($"The Given URL is: {URL}");
        }
        public async Task EnterAcutisCredentialsAsync(string username, string password)
        {
            if (!Directory.Exists(directoryPath))
            {
                // Create the directory
                Directory.CreateDirectory(directoryPath);
            }

            // Check if the file exists
            if (!File.Exists(path))
            {
                File.Create(path).Dispose(); // Ensure the file is properly closed after creation
            }

            await File.WriteAllTextAsync(path, string.Empty);

            await TxtAcuUserName.FillAsync(username);
            await TxtAcuPassword.FillAsync(password);
            await BtnAcuLogin.ClickAsync();
            await MenuSchool.ClickAsync();
        }

        public async Task OpenSchoolAsync(string searchSchool)
        {
            await TxtAcuSchoolSearch.FillAsync(searchSchool); // Enter the school name

            // Prepare to wait for the new page before clicking
            var waitForNewPage = _page.Context.WaitForPageAsync();

            // Click the NextGen icon (which opens a new tab)
            await IconNextGen.ClickAsync();

            // Wait for the new page to be created
            var newPage = await waitForNewPage;

            // Close the current page
            await _page.CloseAsync();

            // Update _page to reference the new page
            _page = newPage;

            // Bring the new page to front
            await newPage.BringToFrontAsync();
            await VerifyAdminDashboardAsync(); // Verify that the admin dashboard is displayed
        }

        // To navigate back to the Acutis domain
        public async Task BackToAcutisAsync()
        {
            // Navigate to the Acutis school details page
            await _page.GotoAsync("https://acutis.optionc.com/school-details");
        }

        public async Task<bool> VerifyAdminDashboardAsync()
        {
            try
            {
                // Wait for the admin dashboard section to be visible
                var dashboardLocator = _page.Locator("//*[@id='sec-slider']/div");
                await dashboardLocator.WaitForAsync();
                // Check if the admin dashboard element is enabled
                bool isDashboardEnabled = await AdminDashboard.IsEnabledAsync();

                if (!isDashboardEnabled)
                {
                    Console.WriteLine("Admin Dashboard is NOT enabled.");
                    return false;
                }
                Console.WriteLine("Logged in Successfully and Admin Dashboard is opened.");              
                await SchoolName.WaitForAsync(); // Wait for the school name element to be present          
                string getSchoolName = (await SchoolName.TextContentAsync())?.Trim() ?? "Unknown School";// Get and trim the school name text

                // Ensure the log file directory exists
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }
                await File.AppendAllTextAsync(path, $"School Name: {getSchoolName}\n");// Append the school name to the audit log file
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in VerifyAdminDashboardAsync: {ex.Message}");
                return false;
            }
        }
        // This method navigates to the Fee Management section
        public async Task NavigateToFeeManagementAsync()
        {
            // Wait for the Administration menu to be visible and click it
            await AdministrationMenu.WaitForAsync();
            await AdministrationMenu.ClickAsync();

            // Wait for the Fee Management menu to be visible and click it
            await FeeManagementMenu.WaitForAsync();
            await FeeManagementMenu.ClickAsync();
        }
        public async Task VerifyFeeManagementAsync()
        {
            // Wait for the Fee Reports page to be visible
            await FeeReports.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });

            // Ensure the element is enabled before interacting
            if (!await FeeReports.IsEnabledAsync())
            {
                throw new Exception("Fee Reports page is not enabled.");
            }

            Console.WriteLine("Fee Reports page is opened");

            // Click on the Fee Reports page
            await FeeReports.ClickAsync();
        }

        // This method verifies whether the Fee Management dashboard is displayed
        public async Task<(bool, long)> SearchFeeReportAsync(long reportId, IPage page)
        {
            // Wait for the search box
            await FeeReportsSearch.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });

            if (!await FeeReportsSearch.IsEnabledAsync())
            {
                Console.WriteLine("Fee reports search box is not enabled.");
                return (false, reportId);
            }

            await FeeReportsSearch.FillAsync(reportId.ToString());

            await FeeReportsFirst.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });

            if (!await FeeReportsFirst.IsEnabledAsync())
            {
                Console.WriteLine("Fee report result is not enabled.");
                return (false, reportId);
            }

            Console.WriteLine($"{reportId} report is searched");
            await FeeReportsFirst.ClickAsync();
            await Task.Delay(1000);

            return (true, reportId);
        }
        public async Task ChangePageLengthAsync()
        {
            // Wait for the dropdown to be available
            await RecordLength.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });

            // Select the value "-1"
            await RecordLength.SelectOptionAsync(new[] { "-1" });

            Console.WriteLine("Page length changed successfully.");
        }
        public async Task<Report> FundingTransactionAmountAsync(Report objreport)
        {
            try
            {
                // Scroll to the bottom of the page
                await Page.EvaluateAsync("window.scrollBy(0, document.body.scrollHeight)");

                // Wait for the Funding Amount label and get its text
                await LblFundingAmount.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });
                string getFundingAmt = await LblFundingAmount.InnerTextAsync() ?? "0";

                // Wait for the Transaction Amount label and get its text
                await LblTransactionAmount.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });
                string getTransactionAmt = await LblTransactionAmount.InnerTextAsync() ?? "0";

                // Compare the amounts
                string comparisonResults = getFundingAmt == getTransactionAmt
                    ? $"Funding Amount: {getFundingAmt} ; and Transaction Amount: {getTransactionAmt} are MATCHING."
                    : $"Funding Amount: {getFundingAmt} ; and Transaction Amount: {getTransactionAmt} are NOT MATCHING.";

                // Store the results in the report object
                objreport.CompareResults = comparisonResults;

                // Log the result
                Console.WriteLine(comparisonResults);

                return objreport;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while comparing funding and transaction amounts: {ex.Message}");
                throw;
            }
        }

        public async Task<ReportData> StartFilterAsync(Int64 ReportId, Report objReport)
        {
            var reportData = new ReportData();
            string getTransactionAmt = "0.00";

            var pages = Page.Context.Pages; // Get all open pages

            // Close the current page
            await _page.CloseAsync();

            // Switch to the last opened page
            var lastPage = pages.Last();
            await lastPage.BringToFrontAsync();

            if (ReportId != 911)
            {
                // Check if the staff portal link is visible before clicking
                if (await StaffCheck.IsVisibleAsync())
                {
                    await StaffCheck.ClickAsync();
                }
            }

            // Start date filtering
            await TransStartDateFilter.ClickAsync();
            await TransStartDateFilter.FillAsync(DateFunction("01/01/2020"));

            // Click appropriate filter button
            if (ReportId != 911)
                await BtnFilter.ClickAsync();
            else
                await Btn911Filter.ClickAsync();

            // Wait for the table to load
            IReadOnlyList<IElementHandle> tableRows;
            if (ReportId != 911)
            {
                await Page.WaitForSelectorAsync("//*[@id='DataTables_Table_0']", new() { State = WaitForSelectorState.Visible });
                tableRows = await Page.QuerySelectorAllAsync("//tfoot");
            }
            else
            {
                await Table911.WaitForAsync(new() { State = WaitForSelectorState.Visible });
                await Page.WaitForTimeoutAsync(10000);  // Artificial delay (consider replacing with event-based waiting)
                tableRows = await Page.QuerySelectorAllAsync("//tfoot");
            }

            string totalValue = "0.00";
            if (ReportId == 901 || ReportId == 904)
            {
                await RecordLength.SelectOptionAsync("-1");

                var rows = await Page.QuerySelectorAllAsync("//table//tbody/tr");
                decimal amount = 0;
                decimal serviceFeeAmount = 0;

                var rowsCount = await Page.QuerySelectorAllAsync("//td[text()='New']|//td[text()='Processing']");
                if (rowsCount.Count > 0)
                {
                    foreach (var row in rows)
                    {
                        var clientElement = await row.QuerySelectorAsync("td:nth-of-type(9)");
                        if (clientElement != null)
                        {
                            string clientText = await clientElement.InnerTextAsync();
                            if (clientText == "New" || clientText == "Processing")
                            {
                                var amountCell = await row.QuerySelectorAsync("td:nth-of-type(4)");
                                if (amountCell != null)
                                {
                                    string amountText = await amountCell.InnerTextAsync();
                                    amount += !string.IsNullOrEmpty(amountText) ? Convert.ToDecimal(amountText.Replace("$", "")) : 0;

                                    // If amount > 0, calculate service fee (1.5%)
                                    serviceFeeAmount = amount > 0 ? amount * 0.015m : 0;
                                }
                            }
                        }
                    }
                    totalValue = amount.ToString();
                }
                reportData.ReportValue = totalValue;
            }
            else
            {
                totalValue = tableRows.Count > 0
                    ? (ReportId != 911 ? await ReportValue.InnerTextAsync() : await Report911Value.InnerTextAsync())
                    : "0.00";

                if (ReportId == 902 || ReportId == 905)
                {
                    getTransactionAmt = tableRows.Count > 0 ? await LblTransactionAmount.InnerTextAsync() : "0.00";
                    reportData.ComparedResult = (totalValue == getTransactionAmt)
                        ? $"Funded amount {totalValue} and Transaction amount {getTransactionAmt} are equal."
                        : $"Funded amount {totalValue} and Transaction amount {getTransactionAmt} are NOT equal.";
                }
                reportData.ReportValue = totalValue;
            }

            return reportData;
        }
        public async Task<Report> LoopFilterAsync()
        {
            Int64[] reportIds = new Int64[5] { 911, 901, 904, 902, 905 };
            Report objReport = new Report();

            if (reportIds.Length > 0)
            {
                using (StreamWriter writer = new StreamWriter(path, true))
                {
                    await writer.WriteLineAsync("MM REPORTS STATISTICS" + "\r\n" + "=====================");

                    foreach (var reportId in reportIds)
                    {
                        await Page.GotoAsync("https://feemanagement.optionc.com/report-list");

                        await SearchFeeReportAsync(reportId, Page);
                        ReportData reportData = await StartFilterAsync(reportId, objReport);

                        switch (reportId)
                        {
                            case 901:
                                objReport.Report901 = reportData.ReportValue;
                                await writer.WriteLineAsync($"eCheck Transaction Detail (901) Report Value in New/Processing Status: {reportData.ReportValue}");
                                break;
                            case 904:
                                objReport.Report904 = reportData.ReportValue;
                                await writer.WriteLineAsync($"Credit Card Detail (904) Report Value in New/Processing Status: {reportData.ReportValue}");
                                break;
                            case 902:
                                objReport.Report902 = reportData.ReportValue;
                                objReport.CompareResults = reportData.ComparedResult;
                                await writer.WriteLineAsync($"eCheck Funded Transactions (902) Report Value as on Today: {reportData.ReportValue}");
                                await writer.WriteLineAsync($"902 Report Comparison Result: {reportData.ComparedResult}");
                                break;
                            case 905:
                                objReport.Report905 = reportData.ReportValue;
                                objReport.CompareResults += "\n" + reportData.ComparedResult;
                                await writer.WriteLineAsync($"Credit Card Funding (905) Report Value as on Today: {reportData.ReportValue}");
                                await writer.WriteLineAsync($"905 Report Comparison Result: {reportData.ComparedResult}");
                                break;
                            case 911:
                                objReport.Report911 = reportData.ReportValue;
                                await writer.WriteLineAsync($"eCheck and Credit Card Funded Transaction Summary (911) Report Value as on Today: {reportData.ReportValue}");
                                break;
                        }
                    }
                }
            }
            return objReport;
        }
        public async Task FinalViewAsync()
        {
            var filePath = @"C:\path\to\your\file.txt"; // Update with actual path

            // Open Notepad and load the file
            System.Diagnostics.Process.Start("notepad.exe", filePath);

            // Optional delay (if needed for UI synchronization)
            await Task.Delay(5000);
        }
        public async Task<Report> VerifyAccountStatusAsync(Report objreport, IPage Page)
        {
            await Page.GotoAsync("https://feemanagement.optionc.com/report-list");

            await SearchFeeReportAsync(907, Page);

            var pages = Page.Context.Pages;
            if (pages.Count > 1)
            {
                await pages[0].CloseAsync();
                Page = pages.Last();
            }

            await Page.WaitForSelectorAsync("#staffCheck");
            var staffCheck = Page.Locator("#staffCheck");
            var btnRunReport = Page.Locator("#btnRunReport");

            if (await staffCheck.IsVisibleAsync())
            {
                await staffCheck.ClickAsync();
            }

            if (await btnRunReport.IsEnabledAsync())
            {
                await btnRunReport.ClickAsync();
            }

            var recordLength = Page.Locator("#Recordlength");
            await recordLength.SelectOptionAsync("-1");

            var primaryRows = await Page.Locator("//table//tbody//tr").CountAsync();
            var rowsCount = await Page.Locator("//*[@id='tblParentAccountSetup']/tbody/tr/td[@class='dataTables_empty']").CountAsync();

            var totalPrimaryAccounts = (rowsCount != 1) ? primaryRows.ToString() : "0";
            objreport.TotalPrimaryAccount = totalPrimaryAccounts;

            await using (StreamWriter writer = new StreamWriter(path, true))
            {
                await writer.WriteLineAsync($"\r\nACCOUNT SETUP STATISTICS\r\n========================\r\nTotal Primary Accounts: {totalPrimaryAccounts}");
            }

            await recordLength.SelectOptionAsync("-1");

            var accountCreatedRadio = Page.Locator("#AccountCreatedradio");
            if (await accountCreatedRadio.IsVisibleAsync())
            {
                await accountCreatedRadio.ClickAsync();
            }

            if (await btnRunReport.IsEnabledAsync())
            {
                await btnRunReport.ClickAsync();
            }

            var totalAccountsCreated = await Page.Locator("//table//tbody//tr").CountAsync();
            var rowsCount1 = await Page.Locator("//*[@id='tblParentAccountSetup']/tbody/tr/td[@class='dataTables_empty']").CountAsync();

            objreport.TotalAccountCreated = (rowsCount1 != 1) ? totalAccountsCreated.ToString() : "0";

            await using (StreamWriter writer = new StreamWriter(path, true))
            {
                await writer.WriteLineAsync($"Total Accounts Created: {objreport.TotalAccountCreated}");
            }

            DateTime today = DateTime.Now;
            DateTime yesterday = today.AddDays(-1);
            string yesterdayString = yesterday.ToString("MM/dd/yyyy");

            string accountStatus = string.Empty;
            long creditCardCount = 0;
            long eCheckCount = 0;

            if (rowsCount1 != 1)
            {
                var rows = await Page.Locator("//table//tbody//tr").AllAsync();
                foreach (var row in rows)
                {
                    var status = await row.Locator("td:nth-of-type(7)").InnerTextAsync();
                    var dateText = await row.Locator("td:nth-of-type(8)").InnerTextAsync();
                    var typeText = await row.Locator("td:nth-of-type(6)").InnerTextAsync();
                    var nameText = await row.Locator("td:nth-of-type(4)").InnerTextAsync();

                    string dateFormatted = DateTime.Parse(dateText).ToString("MM/dd/yyyy");

                    if ((status == "In Progress" || status == "In Progress Primary") || (!string.IsNullOrEmpty(dateFormatted) && dateFormatted == yesterdayString))
                    {
                        accountStatus = $"Name: {nameText} ; Date: {dateText} ; Status: {status} ; Account Type: {typeText} ;";

                        await using (StreamWriter writer = new StreamWriter(path, true))
                        {
                            await writer.WriteLineAsync(accountStatus);
                        }
                        objreport.AccountStatus += accountStatus + "\n";
                    }

                    if (typeText == "Credit Card")
                        creditCardCount++;
                    else
                        eCheckCount++;
                }
            }

            objreport.CreditCardCount = creditCardCount.ToString();
            objreport.ECheckCount = eCheckCount.ToString();

            await using (StreamWriter writer = new StreamWriter(path, true))
            {
                await writer.WriteLineAsync($"Total CC Accounts Created: {creditCardCount}\r\nTotal eCheck Accounts Created: {eCheckCount}");
                await writer.WriteLineAsync("###############################################################################");
            }

            return objreport;
        }


        public async Task WriteAuditReportInExcel(
        Report objReport,
        LedgerPayments ledgerPayments = null,
        DashboardAmount dashboardAmount = null,
        int tab = 0,
        string startDate = "",
        long count = 0)
        {
            string getSchoolName = (await SchoolName.TextContentAsync())?.Trim() ?? "Unknown School";
            long finalRow = tab + 4;
            string filePath = Path.Combine(directoryPath, "AuditSummary.xlsx");

            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            if (File.Exists(filePath) && tab == 0)
            {
                File.Delete(filePath);
            }

            WorkBook workBook = File.Exists(filePath) ? WorkBook.Load(filePath) : WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet workSheet = workBook.GetWorkSheet("2024-2025") ?? workBook.CreateWorkSheet("2024-2025");
            


            int lastRow = tab + 2;

            // Styling
            workSheet[$"A2:T{finalRow}"].Style.Font.SetColor(Color.Purple);
            workSheet[$"A2:T{finalRow}"].Style.SetBackgroundColor(Color.Lavender);

            // Assigning values
            workSheet[$"A{lastRow}"].Value = tab + 1;
            workSheet[$"B{lastRow}"].Value = getSchoolName;
            workSheet[$"C{lastRow}"].Value = startDate;
            workSheet[$"D{lastRow}"].Value = dashboardAmount?.CCTotalAmount;
            workSheet[$"E{lastRow}"].Value = dashboardAmount?.eCheckTotalAmount;
            workSheet[$"F{lastRow}"].Value = ConvertToDecimal(dashboardAmount?.TotalAmount ?? "0");
            workSheet[$"G{lastRow}"].Value = ConvertToDecimal(ledgerPayments?.TotalCharges ?? "0");
            workSheet[$"H{lastRow}"].Value = ConvertToDecimal(ledgerPayments?.TotalPayment ?? "0"   );
            workSheet[$"I{lastRow}"].Value = ConvertToDecimal(objReport?.Report904 ?? "0"   );
            workSheet[$"J{lastRow}"].Value = ConvertToDecimal(objReport?.Report901 ?? "0");
            workSheet[$"K{lastRow}"].Value = Math.Round(ConvertToDecimal(objReport?.Report901 ?? "0") * 1.5m / 100, 2);
            workSheet[$"L{lastRow}"].Value = ConvertToDecimal(objReport?.Report905 ?? "0");
            workSheet[$"M{lastRow}"].Value = ConvertToDecimal(objReport?.Report902 ?? "0");
            workSheet[$"N{lastRow}"].Value = ConvertToDecimal(objReport?.Report911 ?? "0");
            workSheet[$"O{lastRow}"].Value = ConvertToInt(objReport?.TotalPrimaryAccount ?? "0");
            workSheet[$"P{lastRow}"].Value = ConvertToInt(objReport?.TotalAccountCreated ?? "0");
            workSheet[$"Q{lastRow}"].Value = ConvertToInt(objReport?.CreditCardCount ?? "0");
            workSheet[$"R{lastRow}"].Value = ConvertToInt(objReport?.ECheckCount ?? "0");
            workSheet[$"S{lastRow}"].Value = objReport?.AccountStatus;
            workSheet[$"T{lastRow}"].Value = objReport?.CompareResults;

            // Conditional Formatting for Comparison Results
            workSheet[$"T{lastRow}"].Style.Font.SetColor(
                objReport?.CompareResults?.Contains("are not equal") == true ? Color.Red : Color.Green);

            // Summary Rows
            workSheet[$"C{finalRow - 1}"].Value = "Sub Total: ";
            workSheet[$"C{finalRow - 1}"].Style.HorizontalAlignment = HorizontalAlignment.Right;

            foreach (char column in "DEFGHIJKLMN")
            {
                string cellAddress = $"{column}{finalRow - 1}";
                workSheet[cellAddress].Value = $"=SUM({column}2:{column}{finalRow - 2})"; // Set formula as a string
            }

            workSheet[$"C{finalRow}"].Value = "Grand Total: ";
            workSheet[$"C{finalRow}"].Style.HorizontalAlignment = HorizontalAlignment.Right;
            workSheet[$"D{finalRow}"].Value = $"=F{finalRow - 1}";

            // Formatting
            workSheet[$"A1:T1"].Style.WrapText = true;
            workSheet[$"R2:T{finalRow}"].Style.WrapText = true;
            workSheet[$"A1:T{finalRow}"].Style.VerticalAlignment = VerticalAlignment.Center;
            workSheet.DisplayGridlines = false;

            // Applying Borders
            for (char column = 'A'; column <= 'T'; column++)
            {
                string range = $"{column}1:{column}{finalRow}";
                workSheet[range].Style.BottomBorder.SetColor(Color.Black);
                workSheet[range].Style.BottomBorder.Type = BorderType.Thin;
                workSheet[range].Style.TopBorder.SetColor(Color.Black);
                workSheet[range].Style.TopBorder.Type = BorderType.Thin;
                workSheet[range].Style.LeftBorder.SetColor(Color.Black);
                workSheet[range].Style.LeftBorder.Type = BorderType.Thin;
                workSheet[range].Style.RightBorder.SetColor(Color.Black);
                workSheet[range].Style.RightBorder.Type = BorderType.Thin;
            }

            workBook.SaveAs(filePath);
            await FinalViewAsync();
        }

        // Helper Methods for Safe Conversion
        private decimal ConvertToDecimal(string value)
        {
            return decimal.TryParse(value?.Replace("$", ""), out decimal result) ? result : 0m;
        }

        private int ConvertToInt(string value)
        {
            return int.TryParse(value, out int result) ? result : 0;
        }



        // To find the total charges and payments posted for the given date range in School Ledger
        public async Task<LedgerPayments> GetChargesPaymentsAsync()
        {
            LedgerPayments ledgerPayments = new LedgerPayments();

            try
            {
                // Wait for the alumni disabled user checkbox and click it
                await ChkAlumniDisableduser.WaitForAsync();
                await ChkAlumniDisableduser.ClickAsync();

                await Task.Delay(5000); // Explicit wait

                string note = "Testing";

                // Enter and clear note text
                await TxtNote.WaitForAsync();
                await TxtNote.FillAsync(note);
                await TxtNote.ClearAsync();

                // Enter and clear reference text
                await TxtReference.WaitForAsync();
                await TxtReference.FillAsync(note);
                await TxtReference.ClearAsync();

                // Get Yesterday's Date
                DateTime yesterday = DateTime.Now.AddDays(-1);
                string startDate = yesterday.ToString("MM/dd/yyyy");

                // Enter start date
                await TxtStartDate.WaitForAsync();
                await TxtStartDate.ClearAsync();
                await Task.Delay(1000);
                await TxtStartDate.FillAsync(startDate);
                await Task.Delay(2000);

                // Check if the school ledger has family or user drop-down
                bool isFamily = await Page.Locator("#familydiv").IsVisibleAsync();

                if (isFamily)
                {
                    await DdlFamilyUser.WaitForAsync();
                    await DdlFamilyUser.ClickAsync();
                    await SelectAll.WaitForAsync();
                    await SelectAll.ClickAsync();
                }
                else
                {
                    await DdlUser.WaitForAsync();
                    await DdlUser.ClickAsync();
                    await SelectAllUsers.WaitForAsync();
                    await SelectAllUsers.ClickAsync();
                }

                // Click on the Ledger Filter button
                await BtnLedgerFilter.WaitForAsync();
                await BtnLedgerFilter.ClickAsync();

                await Task.Delay(4000); // Wait for the table to load

                // Wait for the ledger table to be present
                var ledgerTable = await Page.Locator("#tblLedger").IsVisibleAsync();
                if (!ledgerTable)
                {
                    Console.WriteLine("Ledger table not found.");
                    return ledgerPayments;
                }

                var rows = await Page.Locator("#tblLedger tr").CountAsync();
                string getDebit = "0";

                if (rows > 2)
                {
                    // Get the total charges and payments
                    string getCredit = await LblTotalCredit.InnerTextAsync() ?? "0";
                    getDebit = await LblTotalDebit.InnerTextAsync() ?? "0";

                    getCredit = getCredit.Trim();
                    getDebit = getDebit.Trim();

                    // Log School Ledger Statistics
                    await File.AppendAllTextAsync(path,
                        "SCHOOL LEDGER STATISTICS\n=======================\n" +
                        $"Total Charges Posted in Ledger (Last 24 hrs): {getDebit}\n" +
                        $"Total Credit in Ledger (Last 24 hrs): {getCredit}\n");
                }
                else
                {
                    await File.AppendAllTextAsync(path,
                        "SCHOOL LEDGER STATISTICS\n=======================\n" +
                        $"School Ledger: No Charges or Payments posted for today ({startDate})\n");
                }

                ledgerPayments = await TodaysMMPaymentsAsync(getDebit);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while retrieving School Ledger Payments: {ex.Message}");
            }

            return ledgerPayments;
        }

        // Method to verify whether an element is present in the ledger table
        public static async Task<bool> IsElementPresentInLedger(IPage page, string selector)
        {
            try
            {
                var element = await page.QuerySelectorAsync(selector);
                return element != null;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public async Task<ACutisDashboardAmount> VerifyAcutisMMDashboardAsync()
        {
            var acutisDashboardAmount = new ACutisDashboardAmount();

            // Click on MM Dashboard
            await _page.Locator($"#{EnumCommandAcutis.MenuControlId.liMMDashboard}").ClickAsync();
            await TabMMCustom.ClickAsync();

            // Set Start Date
            await _page.Locator($"#{EnumCommandAcutis.ControlId.txtStartDate}").ClearAsync();
            await _page.Locator($"#{EnumCommandAcutis.ControlId.txtStartDate}").FillAsync("07/01/2023");

            // Set End Date to today’s date
            string endDate = DateTime.Now.ToString("MM/dd/yyyy");
            await _page.Locator($"#{EnumCommandAcutis.ControlId.txtEndDate}").ClearAsync();
            await _page.Locator($"#{EnumCommandAcutis.ControlId.txtEndDate}").FillAsync(endDate);

            // Click Filter Button
            await _page.Locator($"#{EnumCommandAcutis.ControlId.btnpost}").ClickAsync();

            // Enter Search Text
            await TxtAcutisSearch.FillAsync("OptionC Inc.");

            // Fetching values using the reusable GetAmountAsync method
            string CCTotAmt = await GetAmountAsync("#txtAcCreditCardAmount");
            string CCCount = await GetAmountAsync("#txtAcCreditCardCount");
            string CCServiceFee = await GetAmountAsync("#txtAcCCServiceFeeAmount");
            string CCServiceFeeCount = await GetAmountAsync("#txtAcCCServiceFeeCount");
            string eCheckAmt = await GetAmountAsync("#txtAceCheckAmount");
            string eCheckCount = await GetAmountAsync("#txtAceCheckCount");
            string eCheckServiceFee = await GetAmountAsync("#txtAceCheckServiceFeeAmount");
            string eCheckServiceFeeCount = await GetAmountAsync("#txtAceCheckServiceFeeCount");

            // Calculate total amount
            decimal TotalAmount = Convert.ToDecimal(CCTotAmt) + Convert.ToDecimal(eCheckAmt);

            // Prepare summary
            string summary = $@"
            Acutis MM Dashboard Statistics
            ===============================
            Credit Card Amount: {CCTotAmt}
            Credit Card Count: {CCCount}
            Credit Card Service Fee Amount: {CCServiceFee}
            Credit Card Service Fee Count: {CCServiceFeeCount}
            eCheck Amount: {eCheckAmt}
            eCheck Count: {eCheckCount}
            eCheck Service Fee Amount: {eCheckServiceFee}
            eCheck Service Fee Count: {eCheckServiceFeeCount}
            ";

            // Write to file asynchronously
            await File.AppendAllTextAsync(path, summary + Environment.NewLine);

            // Assign values to the object
            acutisDashboardAmount.ccAmount = Convert.ToDecimal(CCTotAmt);
            acutisDashboardAmount.ccCount = Convert.ToDecimal(CCCount);
            acutisDashboardAmount.ccServiceFeeAmount = Convert.ToDecimal(CCServiceFee);
            acutisDashboardAmount.ccServiceFeeCount = Convert.ToDecimal(CCServiceFeeCount);
            acutisDashboardAmount.eCheckAmount = Convert.ToDecimal(eCheckAmt);
            acutisDashboardAmount.eCheckCount = Convert.ToDecimal(eCheckCount);
            acutisDashboardAmount.eCheckServiceFeeAmount = Convert.ToDecimal(eCheckServiceFee);
            acutisDashboardAmount.eCheckServiceFeeCount = Convert.ToDecimal(eCheckServiceFeeCount);

            return acutisDashboardAmount;
        }

        // Reusable method to get amount from UI elements
        private async Task<string> GetAmountAsync(string selector)
        {
            var element = _page.Locator(selector);
            if (await element.CountAsync() > 0)
            {
                string text = await element.InnerTextAsync();
                return string.IsNullOrEmpty(text) ? "0" : text.Replace("$", "").Trim();
            }
            return "0";
        }


        // Method to get today's MM Payments
        public async Task<LedgerPayments> TodaysMMPaymentsAsync(string getDebit)
        {
            var ledgerPayments = new LedgerPayments();
            ledgerPayments.TotalCharges = string.IsNullOrEmpty(getDebit) ? "0" : getDebit;

            var table = await _page.QuerySelectorAsync("#tblLedger"); // Find the ledger table
            if (table == null) return ledgerPayments; // Return if table is not found

            var rows = await _page.QuerySelectorAllAsync("#tblLedger tr"); // Get all table rows
            int rowCount = rows.Count;

            string ccAmount = "0";
            string eCheckAmount = "0";
            decimal totalMMPayment = 0;

            if (rowCount > 2)
            {
                await SelectPaymentMethodAsync("#ddlPaymentMethod", "#selectCreditCard", "#btnLedgerFilter");
                ccAmount = await GetLedgerAmountAsync("#tblLedger_wrapper div table tfoot tr td", "#lblTotalCredit");

                await SelectPaymentMethodAsync("#ddlPaymentMethod", "#selecteCheck", "#btnLedgerFilter");
                eCheckAmount = await GetLedgerAmountAsync("#tblLedger_wrapper div table tfoot tr td", "#lblTotalCredit");

                totalMMPayment = (string.IsNullOrEmpty(ccAmount) ? 0 : Convert.ToDecimal(ccAmount.Replace("$", ""))) +
                                 (string.IsNullOrEmpty(eCheckAmount) ? 0 : Convert.ToDecimal(eCheckAmount.Replace("$", "")));
            }

            await LogPaymentsAsync(ccAmount, eCheckAmount, totalMMPayment);

            ledgerPayments.CCPayment = ccAmount;
            ledgerPayments.eCheckPayment = eCheckAmount;
            ledgerPayments.TotalPayment = totalMMPayment.ToString();

            return ledgerPayments;
        }
        //Write all the values in the excel file for Acutis MM Dashboard
        public async Task WriteAcutisDashboardValuesinExcelAsync(ACutisDashboardAmount acutisDashboardAmt = null, int tab = 0) //tab = row count
        {
            int lastRow = tab + 2;
            //workSheet[$"A{lastRow}"].Value = "";
            await Task.CompletedTask;
        }

        // Selects a payment method and filters results
        private async Task SelectPaymentMethodAsync(string ddlSelector, string optionSelector, string filterButtonSelector)
        {
            await _page.ClickAsync(ddlSelector);
            await _page.ClickAsync(optionSelector);
            await _page.ClickAsync(filterButtonSelector);
            await _page.WaitForTimeoutAsync(3000);
        }

        // Retrieves the payment amount from the ledger
        private async Task<string> GetLedgerAmountAsync(string footerSelector, string labelSelector)
        {
            var isLedgerRowExists = await _page.QuerySelectorAsync(footerSelector) != null;
            if (isLedgerRowExists)
            {
                var labelElement = await _page.QuerySelectorAsync(labelSelector);
                return labelElement != null ? (await labelElement.InnerTextAsync()).Trim() : "0";
            }
            return "0.00";
        }



        // Logs payment details into a file
        private async Task LogPaymentsAsync(string ccAmount, string eCheckAmount, decimal totalMMPayment)
        {
            string logEntry = $"Total Credit Card Payments(Last 24 hrs): {ccAmount}\r\n" +
                              $"Total eCheck Payments(Last 24 hrs): {eCheckAmount}\r\n" +
                              $"Total MM Payment(Last 24 hrs): {totalMMPayment}\r\n";

            await File.AppendAllTextAsync(path, logEntry);
        }

        public async Task<DashboardAmount> VerifySchoolMMDashboardAsync(string startDate)
        {
            DashboardAmount dashboardAmount = new DashboardAmount();

            await _page.Locator("#tabMMDashboard").WaitForAsync();
            await _page.EvaluateAsync("window.scrollTo(0,0);");
            await _page.Locator("#tabMMDashboard").ClickAsync();
            await _page.Locator("#tabCustom").WaitForAsync();
            await _page.Locator("#tabCustom").ClickAsync();

            DateTime today = DateTime.Now;
            string endDate = today.ToString("MM/dd/yyyy");

            await _page.Locator("#txtStartDate").WaitForAsync();
            await _page.Locator("#txtStartDate").FillAsync(DateFunction(startDate));

            await _page.Locator("#txtEndDate").WaitForAsync();
            await _page.Locator("#txtEndDate").FillAsync(DateFunction(endDate));

            await _page.Locator("#btnCustomFilter").WaitForAsync();
            await _page.Locator("#btnCustomFilter").ClickAsync();

            await _page.WaitForTimeoutAsync(1000); // Instead of Thread.Sleep

            string CCTotAmt = "0";
            string eCheckTotAmt = "0";
            decimal TotalAmount = 0;

            await _page.EvaluateAsync("window.scrollBy(0, document.body.scrollHeight);");

            try
            {
                var lblCCAmount = await _page.Locator("#lblCCAmount").InnerTextAsync();
                var lblCCCount = await _page.Locator("#lblCCCount").InnerTextAsync();
                var lbleCheckAmount = await _page.Locator("#lbleCheckAmount").InnerTextAsync();
                var lbleCheckCount = await _page.Locator("#lbleCheckCount").InnerTextAsync();

                string getCCAmt = string.IsNullOrEmpty(lblCCAmount) ? "0" : lblCCAmount.Replace("$", "");
                CCTotAmt = string.IsNullOrEmpty(getCCAmt) ? "0" : getCCAmt;

                string geteCheckAmt = string.IsNullOrEmpty(lbleCheckAmount) ? "0" : lbleCheckAmount.Replace("$", "");
                eCheckTotAmt = string.IsNullOrEmpty(geteCheckAmt) ? "0" : geteCheckAmt;

                TotalAmount = (string.IsNullOrEmpty(CCTotAmt) ? 0 : Convert.ToDecimal(CCTotAmt))
                            + (string.IsNullOrEmpty(eCheckTotAmt) ? 0 : Convert.ToDecimal(eCheckTotAmt));

                string value = $"MATT MONEY DASHBOARD STATISTICS\n" +
                               $"===============================\n" +
                               $"Credit Card Amount: {CCTotAmt}\n" +
                               $"Credit Card Count: {lblCCCount}\n" +
                               $"eCheck Amount: {eCheckTotAmt}\n" +
                               $"eCheck Count: {lbleCheckCount}\n" +
                               $"Total Amount for current year: {TotalAmount}\n";

                await _page.EvaluateAsync("window.scrollTo(0,0);");
                await File.AppendAllTextAsync(path, value);
            }
            catch (Exception)
            {
                await _page.EvaluateAsync("window.scrollTo(0,0);");
                string errorValue = $"MATT MONEY DASHBOARD STATISTICS\n" +
                                    $"===============================\n" +
                                    $"Credit Card Amount: {CCTotAmt}\n" +
                                    $"eCheck Amount: {eCheckTotAmt}\n" +
                                    $"Total Amount for current year: {TotalAmount}\n";
                await File.AppendAllTextAsync(path, errorValue);
            }

            dashboardAmount.CCTotalAmount = Convert.ToDecimal(CCTotAmt);
            dashboardAmount.eCheckTotalAmount = Convert.ToDecimal(eCheckTotAmt);
            dashboardAmount.TotalAmount = TotalAmount.ToString();

            return dashboardAmount;
        }

        private string DateFunction(string date)
        {
            return date; // Modify if needed
        }
    }

   
}



