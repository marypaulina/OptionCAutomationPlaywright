using OfficeOpenXml;
using Microsoft.Playwright;
using NUnit.Framework;
using OptionCSMSAutomationPlayWright.Model;
using Color = System.Drawing.Color;
using IElementHandle = Microsoft.Playwright.IElementHandle;
using IPage = Microsoft.Playwright.IPage;
using Page = DocumentFormat.OpenXml.Spreadsheet.Page;
using Path = System.IO.Path;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using static OptionCSMSAutomationPlayWright.SISPages.BasePageObject;
using OpenQA.Selenium;
using CsvHelper;
using OpenQA.Selenium.BiDi.Modules.BrowsingContext;
using System.Globalization;

namespace OptionCSMSAutomationPlayWright.SISPages
{
    public class MMAuditPages : BasePageObject
    {
        private IPage _page;
        private WaitHelper waitHelper;
        private readonly string directoryPath = "D:\\File";
        public string path = string.Empty;

        public MMAuditPages(IPage page, WaitHelper waitHelper) : base(page)
        {         
            _page = page;
            _page.SetDefaultTimeout(100000);
            this.waitHelper = waitHelper;
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
        public ILocator FeeManagementMenu => _page.Locator("(//*[@id='li_billing_ledger_view'])[1]");
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

        /* Timeout Methods
         * (new() { Timeout = 60000 });
        await _page.GotoAsync(URL, new() { Timeout = 100000 });
        await _page.GotoAsync("https://acutis.optionc.com/school-details", new() { Timeout = 80000 });
        await Table911.WaitForAsync(new() { State = WaitForSelectorState.Visible });*/

        #endregion

        //Methods for passing the user credentials
        //1. Load URL method    
        public async Task LoadURLAsync(string URL)
        {
            await _page.GotoAsync(URL, new() { Timeout = 100000 });
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


        public async Task MaxWindow()
        {
            // **Ensure the new page is maximized**
            await _page.SetViewportSizeAsync(1500, 1080); // Resize the viewport
            await _page.EvaluateAsync("window.moveTo(0, 0); window.resizeTo(screen.width, screen.height);"); // Maximize using JavaScript
        }

        public async Task OpenSchoolAsync(string searchSchool)
        {
            await TxtAcuSchoolSearch.FillAsync(searchSchool); // Enter the school name    

            var waitForNewPage = _page.Context.WaitForPageAsync(); // Prepare to wait for the new page before clicking

            await IconNextGen.First.ClickAsync(new() { Timeout = 20000 }); // Click the NextGen icon (which opens a new tab)

            var newPage = await waitForNewPage; // Wait for the new page to be created

            if (newPage == null)
            {
                throw new Exception("Failed to open the new page.");
            }

            if (!_page.IsClosed)
            {
                await _page.CloseAsync(); // Close the current page only if it's still open
            }
            _page = newPage; // Update _page to reference the new page

            await MaxWindow();

            await _page.BringToFrontAsync(); // Bring the new page to the front
            await VerifyAdminDashboardAsync(); // Verify that the admin dashboard is displayed
        }

        // To navigate back to the Acutis domain
        public async Task BackToAcutisAsync()
        {
            await MaxWindow();
            await _page.GotoAsync("https://acutis.optionc.com/school-details", new() { Timeout = 100000 });
            // Navigate to the Acutis school details page
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
            //await AdministrationMenu.First.WaitForAsync();
            await AdministrationMenu.First.ClickAsync(new() { Timeout = 100000 });

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
            await Task.Delay(1000);
        }

        // This method verifies whether the Fee Management dashboard is displayed
        public async Task<(bool, long)> SearchFeeReportAsync(long reportId, IPage page)
        {
            // Wait for the search box
            await FeeReportsSearch.WaitForAsync(new LocatorWaitForOptions
            {
                State = WaitForSelectorState.Visible,
                Timeout = 10000 // Timeout in milliseconds (10 seconds)
            });

            await MaxWindow();
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
            await MaxWindow();
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

        #region Verify Funding and Transaction Amount is matching
        public async Task<Report> FundingTransactionAmountAsync(Report objreport)
        {
            try
            {
                // Scroll to the bottom of the page
                await _page.EvaluateAsync("window.scrollBy(0, document.body.scrollHeight)");
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
        #endregion

        #region Verify Account Status Report data #907
        public async Task<Report> VerifyAccountStatusAsync(Report objreport)
        {
            try
            {
                await MaxWindow();
                await _page.GotoAsync("https://feemanagement.optionc.com/report-list", new() { Timeout = 100000 });
                await SearchFeeReportAsync(907, _page);
                await MaxWindow();
                var pages = _page.Context.Pages;
                // Ensure multiple pages exist before switching
                if (pages.Count > 1)
                {
                    var firstPage = pages.FirstOrDefault();
                    if (firstPage != null && !firstPage.IsClosed)
                    {
                        await firstPage.CloseAsync();
                    }
                    // Switch to the last opened page and update _page
                    _page = _page.Context.Pages.LastOrDefault() ?? throw new Exception("Failed to switch to the last opened page.");
                    await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                }
                // Click on StaffCheck if visible
                if (await StaffCheck.IsVisibleAsync())
                {
                    // await StaffCheck.ClickAsync(new() { Timeout = 10000 });
                    // await _page.WaitForTimeoutAsync(30000);

                    await StaffCheck.ClickAsync();
                    await _page.WaitForSelectorAsync("//input[@value='Run Report']", new() { State = WaitForSelectorState.Visible, Timeout = 20000 });

                }
                // Ensure 'Run Report' button is visible before clicking
                // await _page.WaitForSelectorAsync("//input[@value='Run Report']", new() { State = WaitForSelectorState.Visible });
                //if (await _page.Locator("//input[@value='Run Report']").IsVisibleAsync())
                //{
                //    // Go ahead and click
                //    await BtnRunReport.ClickAsync();
                //}

                var runReportLocator = _page.Locator("//input[@value='Run Report']");
                await runReportLocator.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 35000 });

                if (await runReportLocator.IsVisibleAsync())
                {
                    await BtnRunReport.ClickAsync();
                }


                //if (await BtnRunReport.IsEnabledAsync())
                //{
                //    await BtnRunReport.ClickAsync(new() { Timeout = 100000 });
                //    await _page.WaitForTimeoutAsync(100000);
                //    Console.WriteLine("Run Report button clicked.");
                //}
                else
                {
                    throw new Exception("Run Report button is disabled or not interactable.");
                }
                // Wait for dropdown visibility
                await _page.WaitForSelectorAsync("//*[@class='form-control input-xsmall input-inline']", new() { State = WaitForSelectorState.Visible });

                if (await RecordLength.IsVisibleAsync())
                {
                    await RecordLength.SelectOptionAsync("-1");
                    Console.WriteLine("Record Length dropdown set to -1.");
                }
                else
                {
                    throw new Exception("Record Length dropdown not found.");
                }
                await _page.WaitForTimeoutAsync(2000); // Allow table update
                // Fetch row counts
                var primaryRows = await _page.Locator("//table//tbody//tr").CountAsync();
                var rowsCount = await _page.Locator("//*[@id='tblParentAccountSetup']/tbody/tr/td[@class='dataTables_empty']").CountAsync();
                objreport.TotalPrimaryAccount = (rowsCount != 1) ? primaryRows.ToString() : "0";
                await using (StreamWriter writer = new StreamWriter(path, true))
                {
                    await writer.WriteLineAsync($"\r\nACCOUNT SETUP STATISTICS\r\n========================\r\nTotal Primary Accounts: {objreport.TotalPrimaryAccount}");
                }
                // Ensure dropdown selection again
                if (await RecordLength.IsVisibleAsync())
                {
                    await RecordLength.SelectOptionAsync("-1");
                }
                // Click 'Account Created' radio button if visible
                if (await AccountCreatedRadio.IsVisibleAsync())
                {
                    await AccountCreatedRadio.ClickAsync();
                    Console.WriteLine("Account Created Radio button clicked.");
                }
                // Click 'Run Report' again
                if (await BtnRunReport.IsEnabledAsync())
                {
                    await BtnRunReport.ClickAsync(new() { Timeout = 10000 });
                    Console.WriteLine("Run Report clicked again.");
                }
                else
                {
                    Console.WriteLine("Run Report button is disabled.");
                }
                await _page.WaitForTimeoutAsync(2000); // Allow table update
                // Fetch updated row count
                var totalAccountsCreated = await _page.Locator("//table//tbody//tr").CountAsync();
                var rowsCount1 = await _page.Locator("//*[@id='tblParentAccountSetup']/tbody/tr/td[@class='dataTables_empty']").CountAsync();
                objreport.TotalAccountCreated = (rowsCount1 != 1) ? totalAccountsCreated.ToString() : "0";
                await using (StreamWriter writer = new StreamWriter(path, true))
                {
                    await writer.WriteLineAsync($"Total Accounts Created: {objreport.TotalAccountCreated}");
                }
                // Extracting yesterday’s date
                string yesterdayString = DateTime.Now.AddDays(-1).ToString("MM/dd/yyyy");
                long creditCardCount = 0, eCheckCount = 0;
                if (rowsCount1 != 1)
                {
                    var rows = await _page.Locator("//table//tbody//tr").AllAsync();
                    foreach (var row in rows)
                    {
                        var status = await row.Locator("td:nth-of-type(7)").InnerTextAsync();
                        var dateText = await row.Locator("td:nth-of-type(8)").InnerTextAsync();
                        var typeText = await row.Locator("td:nth-of-type(6)").InnerTextAsync();
                        var nameText = await row.Locator("td:nth-of-type(4)").InnerTextAsync();

                        if (DateTime.TryParse(dateText, out DateTime dateParsed))
                        {
                            string dateFormatted = dateParsed.ToString("MM/dd/yyyy");

                            if ((status.Contains("In Progress")) || dateFormatted == yesterdayString)
                            {
                                string accountStatus = $"Name: {nameText} ; Date: {dateText} ; Status: {status} ; Account Type: {typeText} ;";

                                await using (StreamWriter writer = new StreamWriter(path, true))
                                {
                                    await writer.WriteLineAsync(accountStatus);
                                }

                                objreport.AccountStatus += accountStatus + "\n";
                            }

                            if (typeText == "Credit Card") creditCardCount++;
                            else eCheckCount++;
                        }
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
            catch (Exception ex)
            {
                Console.WriteLine($"Error encountered: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Report Button Filter, write the report value and compare results are matching or not
        public async Task<ReportData> StartFilterAsync(Int64 ReportId, Report objReport)
        {

            var reportData = new ReportData();
            string getTransactionAmt = "0.00";
            await Task.Delay(3000);
            // Get all open pages
            var allPages = _page.Context.Pages;

            // Close the current page if it's still open
            if (!_page.IsClosed)
            {
                await _page.CloseAsync();
            }
            // Get the last opened page
            var lastPage = allPages.LastOrDefault();

            if (lastPage == null)
            {
                throw new Exception("Failed to switch to the last opened page.");

            }
            // Update _page to reference the last page
            _page = lastPage;
            await _page.BringToFrontAsync(); // Bring the last page to the front

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
            {
                //await WaitForElementAsync("//input[@onclick='ShowReportDetails(this.id)']");
                await BtnFilter.ClickAsync(new() { Timeout = 20000 });
                await Task.Delay(5000);
            }
            else
            {
                await Btn911Filter.ClickAsync(new() { Timeout = 300000 }); // Click with timeout
                await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 60000 }); // Wait until network requests complete
                await Table911.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 60000 }); // Wait until the table is visible
            }
            // Wait for the table to load
            IReadOnlyList<IElementHandle> tableRows;
            if (ReportId != 911)
            {
               // await _page.WaitForSelectorAsync("//*[@id='DataTables_Table_0']", new() { State = WaitForSelectorState.Visible });
                await _page.WaitForSelectorAsync("//*[@id='DataTables_Table_0']", new() { Timeout = 60000, State = WaitForSelectorState.Visible });

                tableRows = await _page.QuerySelectorAllAsync("//tfoot");
            }
            else
            {
                await Table911.WaitForAsync(new() { State = WaitForSelectorState.Visible });
                await _page.WaitForTimeoutAsync(10000);  // Artificial delay (consider replacing with event-based waiting)
                tableRows = await _page.QuerySelectorAllAsync("//tfoot");
            }
            string totalValue = "0.00";
            if (ReportId == 901 || ReportId == 904)
            {
                await RecordLength.SelectOptionAsync("-1");
                var rows = await _page.QuerySelectorAllAsync("//table//tbody/tr");
                decimal amount = 0;
                decimal serviceFeeAmount = 0;
                var rowsCount = await _page.QuerySelectorAllAsync("//td[text()='New']|//td[text()='Processing']");
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
            //Verify whetehr transaction amount is equal or not equal compared to Report 902 or 905
            else
            {
                totalValue = tableRows.Count > 0 ? (ReportId != 911 ? await ReportValue.First.InnerTextAsync() : await Report911Value.First.InnerTextAsync()) : "0.00";



                /*totalValue = tableRows.Count > 0
                    ? (ReportId != 911 ? await ReportValue.InnerTextAsync() : await Report911Value.InnerTextAsync())
                    : "0.00";*/

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
        #endregion

        #region Loop all the MM Reports, Read and Write amounts
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
                        await MaxWindow();
                        await _page.GotoAsync("https://feemanagement.optionc.com/report-list", new() { Timeout = 100000 });

                        await SearchFeeReportAsync(reportId, _page);
                        await MaxWindow();
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
        #endregion

        #region Write all the data into Excel
        public async Task WriteAuditReportInExcel(
        Report objReport,
        LedgerPayments ledgerPayments = null,
        DashboardAmount dashboardAmount = null,
        int tab = 0,
        string startDate = "",
        long count = 0)
        {
            await SchoolName.WaitForAsync(); // Wait for the school name element to be present
            string getSchoolName = (await SchoolName.TextContentAsync())?.Trim() ?? "Unknown School";
            //long finalRow = tab + 4;
            string filePath = Path.Combine(directoryPath, "AuditSummary.xlsx");
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
            if (File.Exists(filePath) && tab == 0)
            {
                File.Delete(filePath);
            }
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "2024-2025")
                                              ?? package.Workbook.Worksheets.Add("2024-2025");
                // Set the column widths manually
                workSheet.Columns[1].Width = 3;
                workSheet.Columns[2].Width = 35;
                workSheet.Columns[3].Width = 13;
                workSheet.Columns[4].Width = 13;
                workSheet.Columns[5].Width = 13;
                workSheet.Columns[6].Width = 13;
                workSheet.Columns[7].Width = 12;
                workSheet.Columns[8].Width = 12;
                workSheet.Columns[9].Width = 12;
                workSheet.Columns[10].Width = 12;
                workSheet.Columns[11].Width = 10;
                workSheet.Columns[12].Width = 13;
                workSheet.Columns[13].Width = 15;
                workSheet.Columns[14].Width = 15;
                workSheet.Columns[15].Width = 8;
                workSheet.Columns[16].Width = 8;
                workSheet.Columns[17].Width = 10;
                workSheet.Columns[18].Width = 10;
                workSheet.Columns[19].Width = 63;
                workSheet.Columns[20].Width = 65;             
                int lastRow = tab + 2;
                int headerRow = 1;
                int dataStartRow = 2 + tab;
                int finalRow = dataStartRow + 2;
                // Define Column Headers (if it's a new sheet)
                if (workSheet.Dimension == null || workSheet.Dimension.Rows == 0)
                {
                    string[] headers = {
                        "#",
                        "School Name",
                        "Date From",
                        "CC Total Amount (Dashboard)",
                        "eCheck Total Amount (Dashboard)",
                        "Total Amount (Dashboard)",
                        "Last Charges (Last 24 hrs)",
                        "Last Payment using MM (Last 24 hrs)",
                        "CC Amount in New Status (904)",
                        "eCheck Amount in New Status (901)",
                        "Service Fee Amount",
                        "CC Amt Funded (905) as on today",
                        "eCheck Amt Funded (902) as on today",
                        "Tot Amt Funded (911) as on today",
                        "Total Primary Acc.",
                        "Total Acc. Created",
                        "Total CC Acc. Created",
                        "Total eCheck Acc. Created",
                        "New/ InProgress Account",
                        "Funding Amount vs Transaction Amount (902 & 905)"
                    };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = workSheet.Cells[headerRow, i + 1];
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(Color.DarkOliveGreen);
                        cell.Style.Font.Color.SetColor(Color.White);
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                }
                // Assigning values
                workSheet.Cells[$"A{lastRow}"].Value = tab + 1;
                workSheet.Cells[$"B{lastRow}"].Value = getSchoolName;
                workSheet.Cells[$"C{lastRow}"].Value = startDate;
                workSheet.Cells[$"D{lastRow}"].Value = dashboardAmount?.CCTotalAmount;
                workSheet.Cells[$"E{lastRow}"].Value = dashboardAmount?.eCheckTotalAmount;
                workSheet.Cells[$"F{lastRow}"].Value = ConvertToDecimal(dashboardAmount?.TotalAmount ?? "0");
                workSheet.Cells[$"G{lastRow}"].Value = ConvertToDecimal(ledgerPayments?.TotalCharges ?? "0");
                workSheet.Cells[$"H{lastRow}"].Value = ConvertToDecimal(ledgerPayments?.TotalPayment ?? "0");
                workSheet.Cells[$"I{lastRow}"].Value = ConvertToDecimal(objReport?.Report904 ?? "0");
                workSheet.Cells[$"J{lastRow}"].Value = ConvertToDecimal(objReport?.Report901 ?? "0");
                workSheet.Cells[$"K{lastRow}"].Value = Math.Round(ConvertToDecimal(objReport?.Report901 ?? "0") * 1.5m / 100, 2);
                workSheet.Cells[$"L{lastRow}"].Value = ConvertToDecimal(objReport?.Report905 ?? "0");
                workSheet.Cells[$"M{lastRow}"].Value = ConvertToDecimal(objReport?.Report902 ?? "0");
                workSheet.Cells[$"N{lastRow}"].Value = ConvertToDecimal(objReport?.Report911 ?? "0");
                workSheet.Cells[$"O{lastRow}"].Value = ConvertToInt(objReport?.TotalPrimaryAccount ?? "0");
                workSheet.Cells[$"P{lastRow}"].Value = ConvertToInt(objReport?.TotalAccountCreated ?? "0");
                workSheet.Cells[$"Q{lastRow}"].Value = ConvertToInt(objReport?.CreditCardCount ?? "0");
                workSheet.Cells[$"R{lastRow}"].Value = ConvertToInt(objReport?.ECheckCount ?? "0");
                workSheet.Cells[$"S{lastRow}"].Value = objReport?.AccountStatus;
                workSheet.Cells[$"T{lastRow}"].Value = objReport?.CompareResults;
                // Apply Row Formatting
                for (int col = 1; col <= workSheet.Dimension.Columns; col++)
                {
                    var cell = workSheet.Cells[lastRow, col];
                    cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    cell.Style.Font.Size = 11;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(Color.MintCream);
                    cell.Style.Font.Color.SetColor(Color.Black);
                    // Number Formatting for Amount Columns
                    if (col >= 4 && col <= 14) // Columns D to I are financial
                    {
                        cell.Style.Numberformat.Format = "$#,##0.00";
                    }
                }
                // Conditional Formatting for Comparison Results
                workSheet.Cells[$"T{lastRow}"].Style.Font.Color.SetColor(
                    objReport?.CompareResults?.Contains("are not equal") == true ? Color.Red : Color.Green);
                // Summary Rows
                workSheet.Cells[$"C{finalRow - 1}"].Value = "Sub Total: ";
                workSheet.Cells[$"C{finalRow - 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                foreach (char column in "DEFGHIJKLMNOPQR")
                {
                    string cellAddress = $"{column}{finalRow - 1}";
                    workSheet.Cells[cellAddress].Formula = $"SUM({column}2:{column}{finalRow - 2})";

                    if ("OPQR".Contains(column))
                    {
                        // Ensure whole numbers without decimals for OPQR
                        workSheet.Cells[cellAddress].Style.Numberformat.Format = "0";
                    }
                    else
                    {
                        // Apply US currency format ($ and two decimal places) for D to N
                        workSheet.Cells[cellAddress].Style.Numberformat.Format = "$#,##0.00";
                    }
                }
                workSheet.Cells[$"C{finalRow}"].Value = "Grand Total: ";
                workSheet.Cells[$"C{finalRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                workSheet.Cells[$"D{finalRow}"].Formula = $"=F{finalRow - 1}";
                // Formatting
                workSheet.Cells[$"A1:T1"].Style.WrapText = true;
                workSheet.Cells[$"R2:T{finalRow}"].Style.WrapText = true;
                workSheet.Cells[$"A1:T{finalRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells[$"A1:T{finalRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.View.ShowGridLines = false;
                // Applying Borders
                using (ExcelRange range = workSheet.Cells[$"A1:T{finalRow}"])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Top.Color.SetColor(Color.Black);
                    range.Style.Border.Bottom.Color.SetColor(Color.DarkOliveGreen);
                    range.Style.Border.Left.Color.SetColor(Color.Black);
                    range.Style.Border.Right.Color.SetColor(Color.Black);
                }
                package.Save();
            }
        }
        #endregion

        #region Private Conversion Methods and Reusable methods

        // Helper Methods for Safe Conversion
        private decimal ConvertToDecimal(string value)
        {
            return decimal.TryParse(value?.Replace("$", ""), out decimal result) ? result : 0m;
        }

        private int ConvertToInt(string value)
        {
            return int.TryParse(value, out int result) ? result : 0;
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

        private string DateFunction(string date)
        {
            return date; // Modify if needed
        }
        #endregion

        #region Get Charges and Payments from the Fee Management Ledger
        // To find the total charges and payments posted for the given date range in School Ledger
        public async Task<LedgerPayments> GetChargesPaymentsAsync()
        {
            LedgerPayments ledgerPayments = new LedgerPayments();
            // Wait for the alumni disabled user checkbox and click it
            await ChkAlumniDisableduser.WaitForAsync();
            await Task.Delay(2000); // Explicit wait
            await ChkAlumniDisableduser.ClickAsync();
            await Task.Delay(5000); // Explicit wait
            //string note = "Testing";
            // Enter and clear note text
            //await TxtNote.WaitForAsync();
            //await TxtNote.FillAsync(note);
           // await TxtNote.ClearAsync();
            // Enter and clear reference text
            //await TxtReference.WaitForAsync();
           // await TxtReference.FillAsync(note);
            await TxtReference.ClearAsync();

            // Get today's and yesterday's date
            DateTime today = DateTime.Today;
            DateTime yesterday = today.AddDays(-1);
            string startDate = yesterday.ToString("MM/dd/yyyy");

            // Open the calendar by clicking start date input
            await TxtStartDate.ClickAsync();

            // Ensure the calendar is showing the current month
            string expectedMonth = today.ToString("MMMM yyyy");
            var calendarHeader = _page.Locator("//div[contains(@class, 'datepicker-days')]//th[@class='datepicker-switch']");

            while ((await calendarHeader.InnerTextAsync()).Trim() != expectedMonth)
            {
                await _page.Locator("//div[contains(@class, 'datepicker-days')]//th[@class='next']").ClickAsync();
                await Task.Delay(300);
            }

            // Now select yesterday's date cell (based on day number)
          int yesterdayDay = yesterday.Day;
            await _page.Locator($"//div[contains(@class, 'datepicker-days')]//td[contains(@class, 'day') and not(contains(@class,'old')) and not(contains(@class,'new')) and text()='{yesterdayDay}']").ClickAsync();
            await Task.Delay(500);
            await TxtStartDate.PressAsync("Tab");

            // Optional: Log selected date
            string selectedDate = await TxtStartDate.InputValueAsync();
            Console.WriteLine($"Expected Start Date: {startDate}, Selected Start Date: {selectedDate}");



            //Change the date into prvious date daily
            //await _page.Locator($"#{EnumCommandAcutis.ControlId.txtStartDate}").FillAsync("04/08/2025");

            // Check if the school ledger has family or user drop-down
            bool isUser = await _page.Locator("#userdiv").IsVisibleAsync();
            if (isUser)
            {
                await DdlUser.WaitForAsync();
                await DdlUser.ClickAsync();
                await SelectAllUsers.WaitForAsync();
                await SelectAllUsers.ClickAsync();
            }
            else
            {
                await DdlFamilyUser.WaitForAsync();
                await DdlFamilyUser.ClickAsync();
                await SelectAll.WaitForAsync();
                await SelectAll.ClickAsync();
            }
            // Click on the Ledger Filter button
        
            await BtnLedgerFilter.WaitForAsync();
            await BtnLedgerFilter.ClickAsync();
            await Task.Delay(4000); // Wait for the table to load
            // Wait for the ledger table to be present
            var ledgerTable = await _page.Locator("#tblLedger").IsVisibleAsync();
            if (!ledgerTable)
            {
                Console.WriteLine("Ledger table not found.");
                return ledgerPayments;
            }
            var rows = await _page.Locator("#tblLedger tr").CountAsync();
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
            return ledgerPayments;
        }
        #endregion

        #region Some Public common methods
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

        //Write all the values in the excel file for Acutis MM Dashboard
        public async Task WriteAcutisDashboardValuesinExcelAsync(ACutisDashboardAmount acutisDashboardAmt = null, int tab = 0) //tab = row count
        {
            int lastRow = tab + 2;
            //workSheet[$"A{lastRow}"].Value = "";
            await Task.CompletedTask;
        }

        public async Task FinalViewAsync()
        {
            var filePath = @"D:\File\Audit Summary.txt"; // Update with actual path
            //D:\File\Audit Summary.txt

            // Open Notepad and load the file
            System.Diagnostics.Process.Start("notepad.exe", filePath);

            // Optional delay (if needed for UI synchronization)
            await Task.Delay(5000);
        }
        #endregion

        #region Verify Acutis MM Dashboard Amount and return all the dashboard amount
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
        #endregion


        #region Verify Total MM Payments done and read the amount
        // Method to get today's MM Payments
        public async Task<LedgerPayments> TodaysMMPaymentsAsync(string getDebit)
        {
            var ledgerPayments = new LedgerPayments();//Object Creation for Ledger Payment class
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
                await DdlPaymentMethod.WaitForAsync();
                await DdlPaymentMethod.ClickAsync();
                await SelectCreditCard.WaitForAsync();
                await SelectCreditCard.ClickAsync();
                await BtnLedgerFilter.WaitForAsync();
                await BtnLedgerFilter.ClickAsync();
                await Task.Delay(3000);
                bool isLedgerRowsExistCheck = (await _page.QuerySelectorAsync("//*[@id='tblLedger_wrapper']/div[2]/div/div[3]/div/table/tfoot/tr[1]/td[1]")) != null;
                if (isLedgerRowsExistCheck)
                {
                    bool isElementPresent = (await _page.Locator("//*[@id='tblLedger_wrapper']/div[2]/div/div[3]/div/table/tfoot/tr[1]/td[1]").CountAsync()) > 0;
                    // Get the total payments text
                    string getCredit2 = (await LblTotalCredit.InnerTextAsync())?.Trim();
                    getCredit2 = string.IsNullOrEmpty(getCredit2) ? "0" : getCredit2;
                    // Assign to eCheckAmount
                    eCheckAmount = string.IsNullOrEmpty(getCredit2) ? "0" : getCredit2;
                    // Convert and calculate total payment
                    totalMMPayment =
                        (string.IsNullOrEmpty(ccAmount) ? 0 : Convert.ToDecimal(ccAmount.Replace("$", ""))) +
                        (string.IsNullOrEmpty(eCheckAmount) ? 0 : Convert.ToDecimal(eCheckAmount.Replace("$", "")));
                }
                else
                {
                    eCheckAmount = "0.00";
                    totalMMPayment =
                        (string.IsNullOrEmpty(ccAmount) ? 0 : Convert.ToDecimal(ccAmount.Replace("$", ""))) +
                        (string.IsNullOrEmpty(eCheckAmount) ? 0 : Convert.ToDecimal(eCheckAmount.Replace("$", "")));
                }             
            }
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine("Total Credit Card Payments(Last 24 hrs): " + ccAmount + "\r\n" + "Total eCheck Payments(Last 24 hrs): " + eCheckAmount + "\r\n" + "Total MM Payment(Last 24 hrs) : " + totalMMPayment + "\r\n");
            }
            ledgerPayments.CCPayment = ccAmount;
            ledgerPayments.eCheckPayment = eCheckAmount;
            ledgerPayments.TotalPayment = totalMMPayment.ToString();
            return ledgerPayments;
        }
        #endregion


        #region Verify School MM Dashboard and read and write
        public async Task<DashboardAmount> VerifySchoolMMDashboardAsync(string startDate)
        {
            DashboardAmount dashboardAmount = new DashboardAmount();
            await Task.Delay(2000);
            //await _page.Locator("#tabMMDashboard").WaitForAsync();
            await TabMMDashboard.WaitForAsync();
            await _page.EvaluateAsync("window.scrollTo(0,0);");
            //await _page.Locator("#tabMMDashboard").ClickAsync();
            await TabMMDashboard.ClickAsync();
            await TabCustom.WaitForAsync();
            await TabCustom.ClickAsync();           
            DateTime today = DateTime.Now;
            string endDate = today.ToString("MM/dd/yyyy");


            #region Method 1
            //# Method 1 Works fine

            // Parse the passed startDate string
            DateTime dateToSelect = DateTime.ParseExact(startDate, "MM/dd/yyyy", CultureInfo.InvariantCulture);
            int targetDay = dateToSelect.Day;
            int targetMonth = dateToSelect.Month;
            int targetYear = dateToSelect.Year;

            // Open the calendar
            await TxtStartDate.ClickAsync();

            // Navigate calendar to correct month/year
            while (true)
            {
                var displayedMonthYear = await _page.Locator("//div[contains(@class,'datepicker-days')]//th[@class='datepicker-switch']").First.TextContentAsync();
                DateTime displayedDate = DateTime.ParseExact(displayedMonthYear, "MMMM yyyy", CultureInfo.InvariantCulture);

                if (displayedDate.Month == targetMonth && displayedDate.Year == targetYear)
                    break;

                if (displayedDate < dateToSelect)
                {
                    await _page.Locator("//div[contains(@class,'datepicker-days')]//th[@class='next']").First.ClickAsync();
                }
                else
                {
                    await _page.Locator("//div[contains(@class,'datepicker-days')]//th[@class='prev']").First.ClickAsync();
                }

                await Task.Delay(300); // wait briefly for calendar to update
            }

            // Click on the target day
            await _page.Locator($"//div[contains(@class,'datepicker-days')]//td[contains(@class,'day') and not(contains(@class,'old')) and not(contains(@class,'new')) and text()='{targetDay}']").First.ClickAsync();

            #endregion
            // Wait for next element
            await TxtEndDate.WaitForAsync(new() { Timeout = 30000 });

            /*
                        await TxtStartDate.WaitForAsync();
                        await TxtStartDate.FillAsync(DateFunction(startDate), new() { Force = true });
                       // await _page.EvaluateAsync("document.querySelector('#txtStartDate').value = arguments[0];", DateFunction(startDate));

                        await Task.Delay(2000);
                        await TxtStartDate.PressAsync("Tab");
                        await TxtEndDate.WaitForAsync();
                        await TxtEndDate.FillAsync(DateFunction(endDate));*/
            await TxtEndDate.PressAsync("Tab");
            await BtnCustomFilter.WaitForAsync();
            await BtnCustomFilter.ClickAsync();
            await _page.WaitForTimeoutAsync(10000); // Instead of Thread.Sleep
            string CCTotAmt = "0";
            string eCheckTotAmt = "0";
            decimal TotalAmount = 0;
            await _page.EvaluateAsync("window.scrollBy(0, document.body.scrollHeight);");
            try
            {
                var lblCCAmount = await LblCCAmount.InnerTextAsync();           
                var lblCCCount = await LblCCCount.InnerTextAsync();
                var lbleCheckAmount = await LbleCheckAmount.InnerTextAsync();
                var lbleCheckCount = await LbleCheckCount.InnerTextAsync();
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
        #endregion
    }


}



