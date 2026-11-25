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
using OfficeOpenXml;
using System.IO;
using static OptionCSMSAutomationPlayWright.SISPages.BasePageObject;
using OpenQA.Selenium;
using CsvHelper;
using OpenQA.Selenium.BiDi.Modules.BrowsingContext;
using System.Globalization;
using OptionCSMSAutomationPlayWright.Support;

namespace OptionCSMSAutomationPlayWright.SISPages
{
    public class MMAuditPages : BasePageObject
    {
        private const int InitialUiWaitTimeoutMs = 30000;
        private const int MaximumUiWaitTimeoutMs = 180000;
        private const int Report911RangeYears = 1;
        private sealed record Report911RangeResult(
            decimal? Amount,
            int BodyRowCount,
            bool NoRows,
            string? Diagnostics);

        private static readonly HashSet<string> Report911SplitSchoolIdentifiers = new(StringComparer.OrdinalIgnoreCase)
        {
            "5158",
            "St. Joseph Marquette Catholic School",
            "7127",
            "Holy Trinity School - Kewaskum",
            "7291",
            "Cathedral School - Superior",
            "221",
            "Holy Family Classical Academy",
            "6142",
            "St. Francis of Assisi School-Banks",
            "8232",
            "John Paul II Catholic School",
            "8507",
            "Holy Trinity Catholic School - Assumption",
            "8351",
            "Immaculate Conception School - Bellevue",
            "8958",
            "St. Anne Catholic School - Somerset",
            "8400",
            "Piqua Catholic School",
            "142",
            "Immaculate Conception Catholic School",
            "7285",
            "Saint Patrick Catholic School - Hudson"
            
        };

        private IPage _page;
        private WaitHelper waitHelper;
        private readonly string directoryPath = "D:\\File";
        private readonly string auditWorkbookPath;
        public string path = string.Empty;

        public MMAuditPages(IPage page, WaitHelper waitHelper) : base(page)
        {
            _page = page;
            this.waitHelper = waitHelper;
            Directory.CreateDirectory(directoryPath);
            // File names are prefixed with today's India date as MMddyyyy, e.g. "11202026- Audit Summary.txt".
            string datePrefix = GetIndiaToday().ToString("MMddyyyy", CultureInfo.InvariantCulture);
            path = Path.Combine(directoryPath, $"{datePrefix}- Audit Summary.txt");
            auditWorkbookPath = Path.Combine(directoryPath, $"{datePrefix}- AuditSummary.xlsx");
            // Outputs are reset by PrepareAuditOutputs() unless a previous run left failed schools to retry.
        }

        private static DateTime GetIndiaToday()
        {
            DateTime utcNow = DateTime.UtcNow;
            foreach (string id in new[] { "India Standard Time", "Asia/Kolkata" })
            {
                try
                {
                    return TimeZoneInfo.ConvertTimeFromUtc(utcNow, TimeZoneInfo.FindSystemTimeZoneById(id)).Date;
                }
                catch (TimeZoneNotFoundException) { }
                catch (InvalidTimeZoneException) { }
            }
            return utcNow.AddHours(5).AddMinutes(30).Date; // IST is UTC+05:30
        }

        /// <summary>
        /// The school list changes from run to run (1, 2, 15, 30... schools in any order), so rows in the
        /// workbook are keyed by school code, not by list position. Returns the zero-based indexes of the
        /// schools in <paramref name="schoolCodes"/> that still need to be audited today: those that are not in
        /// today's workbook yet or are marked FAILED. If every listed school already succeeded today, all of them
        /// are re-run (their rows are refreshed in place). Existing files are never deleted.
        /// </summary>
        public List<int> PrepareAuditOutputs(IReadOnlyList<string> schoolCodes)
        {
            var codeStatus = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase); // code -> failed?
            if (File.Exists(auditWorkbookPath))
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(auditWorkbookPath)))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == "2026-2027");
                    if (ws?.Dimension != null)
                    {
                        for (int row = 2; row <= ws.Dimension.End.Row; row++)
                        {
                            string code = AuditCodeAt(ws, row);
                            if (code.Length == 0) continue; // totals rows or rows without a school code
                            string status = ws.Cells[$"U{row}"].Text ?? string.Empty;
                            codeStatus[code] = status.StartsWith("FAILED", StringComparison.OrdinalIgnoreCase);
                        }
                    }
                }
            }

            var toRun = new List<int>();
            for (int i = 0; i < schoolCodes.Count; i++)
            {
                string code = schoolCodes[i].Trim();
                if (!codeStatus.TryGetValue(code, out bool failed) || failed) toRun.Add(i);
            }
            if (toRun.Count == 0)
            {
                toRun.AddRange(Enumerable.Range(0, schoolCodes.Count));
            }
            return toRun;
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
        private ILocator Report911ReportTables => _page.Locator(
            "//table[.//thead//th[normalize-space()='Funding Date'] " +
            "and .//tbody/tr and .//tfoot]");
        private ILocator Report911Total => Report911ReportTables.Locator("tfoot td:nth-child(2)");
        private ILocator Report911FooterCells => Report911ReportTables.Locator("tfoot td");
        public ILocator ReportResult => _page.Locator("//*[@id='export-content']");
        public ILocator RecordLength => _page.Locator("//*[@class='form-control input-xsmall input-inline']");
        public ILocator AllRecord => _page.Locator("//*[@class='form-control input-xsmall input-inline']//option[@value='-1']");
        public ILocator AccountCreatedRadio => _page.Locator("//input[@id='primary1']");
        //public ILocator BtnRunReport => _page.Locator("//input[@value='Run Report']");
        public ILocator BtnRunReport => _page.Locator("//input[@type='button' and @value='Run Report' and contains(@class,'btn-info')]");
        public ILocator SchoolName => _page.Locator("//div[@class='oc-school-name']");
        public ILocator LblFundingAmount => _page.Locator("//*[@id='DataTables_Table_0']/tfoot/tr/td[2]/strong");
        public ILocator LblTransactionAmount => _page.Locator("//*[@id='DataTables_Table_0']/tfoot/tr/td[4]/strong");
        // Logout Elements
        public ILocator LinkSignout => _page.Locator("//a[@class='dropdown-item signout']");
        public ILocator RbAccountHistory => _page.Locator("//input[@type='radio' and @id='primary3']");
        public ILocator ThActivatedOn => _page.Locator("//th[@aria-controls='tblParentAccountSetup' and normalize-space()='Activated On']");


        private ILocator ReactivatedPopup => _page.Locator(
            "//div[contains(@class,'modal-content') and contains(@class,'survey-modal-content') and contains(.,'Your MattMoney Account is now in Active')]");
        private ILocator ReactivatedCloseIcon =>
            ReactivatedPopup.Locator(".//a[contains(@class,'uc-under-construction-x')]");

        // ================= LOCATORS (All XPaths Here) =================

        private ILocator DashboardSection => _page.Locator("//*[@id='sec-slider']/div");

        private ILocator MaintenancePopup =>
            _page.Locator(
                "//div[contains(@class,'modal-content') and contains(@class,'survey-modal-content') and contains(.,'Scheduled Unavailability')]");

        private ILocator PopupCloseXButton =>
            MaintenancePopup.Locator(".//a[contains(@class,'uc-under-construction-x')]");

        private ILocator PopupCloseButton =>
            MaintenancePopup.Locator(".//button[normalize-space()='CLOSE']");

      
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
            await _page.GotoAsync(URL, new()
            {
                Timeout = PlaywrightTimeoutPolicy.NavigationTimeoutMs
            });
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

            using (new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.None))
            {
            }
            await TxtAcuUserName.FillAsync(username);
            await TxtAcuPassword.FillAsync(password);
            await BtnAcuLogin.ClickAsync();
            await MenuSchool.ClickAsync();
        }


        public async Task MaxWindow()
        {
            // **Ensure the new page is maximized**
            await _page.SetViewportSizeAsync(1500, 1080); // Resize the viewport
            // The page may still be navigating ("Execution context was destroyed"); wait and retry.
            for (int attempt = 1; attempt <= 4; attempt++)
            {
                try
                {
                    await _page.EvaluateAsync("window.moveTo(0, 0); window.resizeTo(screen.width, screen.height);"); // Maximize using JavaScript
                    return;
                }
                catch (PlaywrightException) when (attempt < 4)
                {
                    try { await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded, new() { Timeout = 15000 }); }
                    catch (Exception) { }
                    await Task.Delay(1500);
                }
                catch (PlaywrightException)
                {
                    // Resizing the window is cosmetic; never fail the audit because of it.
                    Console.WriteLine("Could not resize the window (page kept navigating); continuing.");
                }
            }
        }

        public async Task OpenSchoolAsync(string searchSchool)
        {
            // The Acutis school list can load slowly or not at all (search box missing, or no result row after
            // searching). Retry: reload the school list and search again, up to 3 times.
            const int searchAttempts = 3;
            for (int attempt = 1; attempt <= searchAttempts; attempt++)
            {
                try
                {
                    if (attempt > 1)
                    {
                        Console.WriteLine($"Acutis school search for '{searchSchool}' not ready (attempt {attempt - 1}); reloading the school list.");
                        await _page.GotoAsync(
                            "https://acutis.optionc.com/school-details",
                            new() { Timeout = PlaywrightTimeoutPolicy.NavigationTimeoutMs });
                    }

                    await TxtAcuSchoolSearch.WaitForAsync(
                        new() { State = WaitForSelectorState.Visible, Timeout = 60000 });
                    try { await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 15000 }); }
                    catch (TimeoutException) { }

                    await TxtAcuSchoolSearch.FillAsync(searchSchool); // Enter the school name
                    try { await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 10000 }); }
                    catch (TimeoutException) { }

                    await IconNextGen.First.WaitForAsync(
                        new() { State = WaitForSelectorState.Visible, Timeout = 30000 });
                    break;
                }
                catch (Exception ex) when (attempt < searchAttempts && (ex is TimeoutException || ex is PlaywrightException))
                {
                    Console.WriteLine($"Acutis school search attempt {attempt} failed: {ex.Message.Split('\n')[0]}");
                }
            }

            var waitForNewPage = _page.Context.WaitForPageAsync(); // Prepare to wait for the new page before clicking

            await IconNextGen.First.ClickAsync(new()
            {
                Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
            }); // Click the NextGen icon (which opens a new tab)

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
            PlaywrightTimeoutPolicy.Configure(_page.Context);

            await MaxWindow();

            await _page.BringToFrontAsync(); // Bring the new page to the front
            if (!await VerifyAdminDashboardAsync())
            {
                throw new InvalidOperationException("The admin dashboard did not become ready.");
            }
        }

        // To navigate back to the Acutis domain
        public async Task BackToAcutisAsync()
        {
            // Navigate to the Acutis school details page; retry because the previous school's page may still be
            // navigating or the navigation may be aborted.
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    await MaxWindow();
                    await _page.GotoAsync(
                        "https://acutis.optionc.com/school-details",
                        new() { Timeout = PlaywrightTimeoutPolicy.NavigationTimeoutMs });
                    return;
                }
                catch (Exception ex) when (attempt < 3 && (ex is PlaywrightException || ex is TimeoutException))
                {
                    Console.WriteLine($"Back to Acutis attempt {attempt} failed: {ex.Message.Split('\n')[0]}; retrying.");
                    await Task.Delay(3000);
                }
            }
        }

        public async Task HandleMMMaintenancePopupIfPresentAsync()
        {
            try
            {
                await MaintenancePopup.WaitForAsync(
                    new() { State = WaitForSelectorState.Visible, Timeout = 1000 });
            }
            catch (TimeoutException)
            {
                Console.WriteLine("No MattMoney maintenance popup appeared.");
                return;
            }

            Console.WriteLine("MattMoney Maintenance popup detected.");
            if (await PopupCloseXButton.IsVisibleAsync())
            {
                await PopupCloseXButton.ClickAsync();
            }
            else if (await PopupCloseButton.IsVisibleAsync())
            {
                await PopupCloseButton.ClickAsync();
            }
            else
            {
                throw new InvalidOperationException("The maintenance popup is visible but has no close control.");
            }

            Console.WriteLine("Popup closed successfully.");
        }

        public async Task HandleMMReactivatedPopupIfPresentAsync()
        {
            try
            {
                await ReactivatedPopup.WaitForAsync(
                    new() { State = WaitForSelectorState.Visible, Timeout = 3000 });
            }
            catch (TimeoutException)
            {
                Console.WriteLine("MM Reactivated popup not displayed. Continuing...");
                return;
            }

            Console.WriteLine("MattMoney Reactivated popup detected.");
            await ReactivatedCloseIcon.ClickAsync();
            Console.WriteLine("Reactivated popup closed successfully.");
        }

        //public async Task<bool> VerifyAdminDashboardAsync()
        //    {
        //        try
        //        {
        //            // Wait for the admin dashboard section to be visible
        //            var dashboardLocator = _page.Locator("//*[@id='sec-slider']/div");
        //            await dashboardLocator.WaitForAsync();
        //            // Check if the admin dashboard element is enabled
        //            bool isDashboardEnabled = await AdminDashboard.IsEnabledAsync();
        //            if (!isDashboardEnabled)
        //            {
        //                Console.WriteLine("Admin Dashboard is NOT enabled.");
        //                return false;
        //            }
        //            Console.WriteLine("Logged in Successfully and Admin Dashboard is opened.");
        //            await SchoolName.WaitForAsync(); // Wait for the school name element to be present          
        //            string getSchoolName = (await SchoolName.TextContentAsync())?.Trim() ?? "Unknown School";// Get and trim the school name text
        //            // Ensure the log file directory exists
        //            if (!Directory.Exists(directoryPath))
        //            {
        //                Directory.CreateDirectory(directoryPath);
        //            }
        //            await File.AppendAllTextAsync(path, $"School Name: {getSchoolName}\n");// Append the school name to the audit log file
        //            return true;
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine($"Error in VerifyAdminDashboardAsync: {ex.Message}");
        //            return false;
        //        }
        //    }

        public async Task<bool> VerifyAdminDashboardAsync()
        {
            try
            {             
                // Wait for the admin dashboard section to be visible
                var dashboardLocator = _page.Locator("//*[@id='sec-slider']/div");
                await dashboardLocator.WaitForAsync(new() { State = WaitForSelectorState.Visible });

                // Check if the admin dashboard element is enabled
                bool isDashboardEnabled = await AdminDashboard.IsEnabledAsync();
                if (!isDashboardEnabled)
                {
                    Console.WriteLine("Admin Dashboard is NOT enabled.");
                    return false;
                }

                Console.WriteLine("Logged in Successfully and Admin Dashboard is opened.");

                // Wait for school name
                await SchoolName.WaitForAsync();
                string getSchoolName =
                    (await SchoolName.TextContentAsync())?.Trim() ?? "Unknown School";

                // Ensure log directory exists
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                // Log school name
                await File.AppendAllTextAsync(path, $"School Name: {getSchoolName}\n");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in VerifyAdminDashboardAsync: {ex.Message}");
                throw;
            }
        }

        private async Task WaitForLoadingOverlayToDisappearAsync(
            int timeout = MaximumUiWaitTimeoutMs,
            ILocator? readyLocator = null)
        {
            if (timeout <= 0)
                throw new ArgumentOutOfRangeException(nameof(timeout), "Wait timeout must be greater than zero.");

            var boundedTimeout = Math.Min(timeout, MaximumUiWaitTimeoutMs);
            var initialTimeout = Math.Min(InitialUiWaitTimeoutMs, boundedTimeout);
            var overlay = _page.Locator(".loaderimg_align:visible");

            try
            {
                await overlay.WaitForAsync(
                    new() { State = WaitForSelectorState.Hidden, Timeout = initialTimeout });
            }
            catch (TimeoutException) when (boundedTimeout > initialTimeout)
            {
                var fallbackTimeout = boundedTimeout - initialTimeout;
                Console.WriteLine(
                    $"Loading overlay is still visible after {initialTimeout}ms; " +
                    $"waiting up to {fallbackTimeout}ms for the page to become ready.");

                if (readyLocator != null)
                {
                    try
                    {
                        await readyLocator.WaitForAsync(
                            new() { State = WaitForSelectorState.Visible, Timeout = fallbackTimeout });
                    }
                    catch (TimeoutException ex)
                    {
                        throw new TimeoutException(
                            $"The loading overlay remained visible and the expected page content " +
                            $"did not appear within {boundedTimeout}ms.",
                            ex);
                    }

                    return;
                }

                await overlay.WaitForAsync(
                    new() { State = WaitForSelectorState.Hidden, Timeout = fallbackTimeout });
            }
        }

        private static async Task WaitForLocatorWithFallbackAsync(
            ILocator locator,
            WaitForSelectorState state = WaitForSelectorState.Visible,
            int timeout = MaximumUiWaitTimeoutMs)
        {
            if (timeout <= 0)
                throw new ArgumentOutOfRangeException(nameof(timeout), "Wait timeout must be greater than zero.");

            var boundedTimeout = Math.Min(timeout, MaximumUiWaitTimeoutMs);
            var initialTimeout = Math.Min(InitialUiWaitTimeoutMs, boundedTimeout);

            try
            {
                await locator.WaitForAsync(new() { State = state, Timeout = initialTimeout });
            }
            catch (TimeoutException) when (boundedTimeout > initialTimeout)
            {
                var fallbackTimeout = boundedTimeout - initialTimeout;
                Console.WriteLine(
                    $"Page content was not ready after {initialTimeout}ms; " +
                    $"waiting up to {fallbackTimeout}ms more.");
                await locator.WaitForAsync(new() { State = state, Timeout = fallbackTimeout });
            }
        }

        private static IEnumerable<(DateTime StartDate, DateTime EndDate)> GetReport911DateRanges()
        {
            var rangeStart = new DateTime(2020, 1, 1);
            var finalDate = DateTime.Today;

            while (rangeStart <= finalDate)
            {
                var rangeEnd = rangeStart.AddYears(Report911RangeYears).AddDays(-1);
                if (rangeEnd > finalDate)
                    rangeEnd = finalDate;

                yield return (rangeStart, rangeEnd);
                // Ranges are inclusive; begin the next range the day after this one ends.
                rangeStart = rangeEnd.AddDays(1);
            }
        }

        private async Task EnsureShowStaffCheckedAsync()
        {
            if (await StaffCheck.CountAsync() > 0)
            {
                var staffCheckbox = StaffCheck.First;
                if (await staffCheckbox.IsVisibleAsync() && !await staffCheckbox.IsCheckedAsync())
                {
                    if (!await staffCheckbox.IsEnabledAsync())
                        throw new InvalidOperationException("The Show Staff checkbox is visible but disabled.");

                    await staffCheckbox.CheckAsync();
                    await WaitForLoadingOverlayToDisappearAsync(timeout: 15000);
                }
            }

            await EnsureShowCheckboxesCheckedAsync();
        }

        // Ticks any visible, enabled, unchecked checkbox whose label starts with "Show"
        // (e.g. "Show Staff", "Show Alumni/Disabled users"), whatever its id is.
        private async Task EnsureShowCheckboxesCheckedAsync()
        {
            var boxes = _page.Locator("input[type='checkbox']:visible:not(:checked)");
            int count = await boxes.CountAsync();
            for (int i = 0; i < count; i++)
            {
                var box = boxes.Nth(i);
                try
                {
                    var labelText = await box.EvaluateAsync<string>(
                        @"el => {
                            let t = '';
                            if (el.id) { const l = document.querySelector('label[for=""' + el.id + '""]'); if (l) t = l.innerText; }
                            if (!t) { const p = el.closest('label'); if (p) t = p.innerText; }
                            if (!t && el.nextSibling) t = el.nextSibling.textContent || '';
                            if (!t && el.nextElementSibling) t = el.nextElementSibling.innerText || '';
                            if (!t && el.parentElement) t = el.parentElement.innerText || '';
                            return (t || '').trim();
                        }");
                    if (!labelText.StartsWith("Show", StringComparison.OrdinalIgnoreCase)
                        || !await box.IsEnabledAsync())
                        continue;

                    await box.CheckAsync();
                    await WaitForLoadingOverlayToDisappearAsync(timeout: 15000);
                }
                catch (PlaywrightException)
                {
                    // Page re-rendered while ticking; remaining boxes are re-evaluated on the next call.
                    break;
                }
            }
        }

        private static bool IsReport911NoDataMessage(string text)
        {
            return text.Contains("no data", StringComparison.OrdinalIgnoreCase)
                || text.Contains("no record", StringComparison.OrdinalIgnoreCase)
                || text.Contains("no matching", StringComparison.OrdinalIgnoreCase)
                || text.Contains("no transaction", StringComparison.OrdinalIgnoreCase)
                || text.Contains("no result", StringComparison.OrdinalIgnoreCase);
        }

        private async Task SetDateInputValueAsync(ILocator dateInput, DateTime date, string fieldName)
        {
            await dateInput.WaitForAsync(new() { State = WaitForSelectorState.Visible });

            if (!await dateInput.IsEnabledAsync())
            {
                throw new InvalidOperationException($"{fieldName} date field is disabled.");
            }

            var inputType = await dateInput.GetAttributeAsync("type");
            var expectedValue = string.Equals(inputType, "date", StringComparison.OrdinalIgnoreCase)
                ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : date.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            var actualValue = await dateInput.InputValueAsync();
            if (await dateInput.GetAttributeAsync("readonly") == null)
            {
                await dateInput.FillAsync(expectedValue);
                await dateInput.PressAsync("Tab");
                actualValue = await dateInput.InputValueAsync();
            }

            if (!IsDateValue(date, actualValue))
            {
                var datepickerUpdated = await dateInput.EvaluateAsync<bool>(
                    @"(input, value) => {
                        const jquery = window.jQuery;
                        if (!jquery || !jquery.fn || typeof jquery.fn.datepicker !== 'function')
                            return false;
                        const parts = value.split(/[/-]/).map(Number);
                        const [year, month, day] = parts[0] > 999
                            ? parts
                            : [parts[2], parts[0], parts[1]];
                        if (!year || !month || !day)
                            return false;
                        jquery(input).datepicker('update', new Date(year, month - 1, day));
                        return true;
                    }",
                    expectedValue);
                actualValue = await dateInput.InputValueAsync();
                if (!datepickerUpdated || !IsDateValue(date, actualValue))
                {
                    await dateInput.EvaluateAsync(
                        @"(input, value) => {
                            const valueSetter = Object.getOwnPropertyDescriptor(
                                HTMLInputElement.prototype,
                                'value').set;
                            valueSetter.call(input, value);
                            input.dispatchEvent(new Event('input', { bubbles: true }));
                            input.dispatchEvent(new Event('change', { bubbles: true }));
                        }",
                        expectedValue);
                    actualValue = await dateInput.InputValueAsync();
                }

                if (!IsDateValue(date, actualValue))
                {
                    throw new InvalidOperationException(
                        $"Could not set {fieldName} date to {expectedValue}; " +
                        $"the field contains '{actualValue}'.");
                }
            }
        }

        private Task SetReportDateFilterAsync(ILocator dateInput, DateTime date, string fieldName)
        {
            return SetDateInputValueAsync(dateInput, date, $"Report 911 {fieldName}");
        }

        private async Task OpenFreshReportPageAsync(long reportId)
        {
            var reportListPage = _page;
            await reportListPage.GotoAsync(
                "https://feemanagement.optionc.com/report-list",
                new() { Timeout = PlaywrightTimeoutPolicy.NavigationTimeoutMs });

            var (reportFound, _) = await SearchFeeReportAsync(reportId, reportListPage);
            if (!reportFound)
                throw new InvalidOperationException($"Could not open fee report {reportId} for the next date range.");

            var reportPage = reportListPage.Context.Pages.LastOrDefault(
                page => !ReferenceEquals(page, reportListPage) && !page.IsClosed);
            if (reportPage == null)
                throw new InvalidOperationException($"Opening fee report {reportId} did not create a report page.");

            await reportListPage.CloseAsync();
            _page = reportPage;
            PlaywrightTimeoutPolicy.Configure(_page.Context);
            await _page.BringToFrontAsync();
            await MaxWindow();
        }

        private async Task<decimal> GetReport911TotalAsync(StreamWriter auditWriter, string schoolIdentifier)
        {
            await TransStartDateFilter.WaitForAsync(new() { State = WaitForSelectorState.Visible });
            await Btn911Filter.WaitForAsync(new() { State = WaitForSelectorState.Visible });

            var endDateFilter = TransEndDateFilter.First;
            var hasEditableEndDateFilter = await TransEndDateFilter.CountAsync() > 0
                && await endDateFilter.IsVisibleAsync()
                && await endDateFilter.IsEnabledAsync()
                && await endDateFilter.GetAttributeAsync("readonly") == null;

            var shouldSplitDateRange = hasEditableEndDateFilter
                && Report911SplitSchoolIdentifiers.Contains(schoolIdentifier.Trim());
            var firstDate = new DateTime(2020, 1, 1);
            var ranges = shouldSplitDateRange
                ? GetReport911DateRanges().ToArray()
                : new[] { (StartDate: firstDate, EndDate: DateTime.Today) };
            var rangeDescription = shouldSplitDateRange
                ? string.Empty
                : hasEditableEndDateFilter
                    ? " (single query)"
                    : " (single query; no end-date filter)";
            if (ranges.Length == 0)
                throw new InvalidOperationException("No Report 911 date ranges were generated.");

            if (!shouldSplitDateRange)
            {
                var reason = hasEditableEndDateFilter
                    ? $"school {schoolIdentifier} is not configured for split ranges"
                    : "there is no editable end-date filter";
                await auditWriter.WriteLineAsync(
                    $"Report 911 {reason}; running one unsplit query from " +
                    $"{firstDate:MM/dd/yyyy} through {DateTime.Today:MM/dd/yyyy}.");
                await auditWriter.FlushAsync();
            }

            decimal total = 0m;
            for (var index = 0; index < ranges.Length; index++)
            {
                var (startDate, endDate) = ranges[index];
                if (index > 0)
                {
                    // Stay on the Report 911 page and just change the date range filter.
                    // Only go back to the report manager if the filters are not available on this page.
                    var filtersAvailable = false;
                    try
                    {
                        await TransStartDateFilter.WaitForAsync(
                            new() { State = WaitForSelectorState.Visible, Timeout = 10000 });
                        await Btn911Filter.WaitForAsync(
                            new() { State = WaitForSelectorState.Visible, Timeout = 10000 });
                        filtersAvailable = true;
                    }
                    catch (TimeoutException)
                    {
                    }

                    if (!filtersAvailable)
                    {
                        await auditWriter.WriteLineAsync(
                            "Report 911 date filters not available on the result page; reopening the report.");
                        await auditWriter.FlushAsync();
                        await OpenFreshReportPageAsync(911);
                    }
                }

                var rangeResult = await RunReport911RangeAsync(
                    startDate,
                    endDate,
                    hasEditableEndDateFilter);
                if (rangeResult.Amount == null && !rangeResult.NoRows)
                {
                    await auditWriter.WriteLineAsync(
                        $"Report 911 did not show a total for {startDate:MM/dd/yyyy} - " +
                        $"{endDate:MM/dd/yyyy}; retrying once on a fresh report page.");
                    await auditWriter.FlushAsync();
                    await OpenFreshReportPageAsync(911);
                    rangeResult = await RunReport911RangeAsync(
                        startDate,
                        endDate,
                        hasEditableEndDateFilter);
                }

                if (rangeResult.Amount == null)
                {
                    throw new InvalidOperationException(
                        $"Report 911 returned {rangeResult.BodyRowCount} rows for " +
                        $"{startDate:MM/dd/yyyy} - {endDate:MM/dd/yyyy} but no total appeared " +
                        $"after retry. Table diagnostics: {rangeResult.Diagnostics}");
                }

                if (rangeResult.NoRows)
                {
                    await auditWriter.WriteLineAsync(
                        $"Report 911 {startDate:MM/dd/yyyy} - {endDate:MM/dd/yyyy}" +
                        $"{rangeDescription}: $0.00 (no rows)");
                    await auditWriter.FlushAsync();
                    continue;
                }

                var rangeAmount = rangeResult.Amount.Value;
                total += rangeAmount;
                await auditWriter.WriteLineAsync(
                    $"Report 911 {startDate:MM/dd/yyyy} - {endDate:MM/dd/yyyy}" +
                    $"{rangeDescription}: " +
                    rangeAmount.ToString("C2", CultureInfo.GetCultureInfo("en-US")));
                await auditWriter.FlushAsync();
            }

            await auditWriter.WriteLineAsync(
                $"Report 911 combined total: {total.ToString("C2", CultureInfo.GetCultureInfo("en-US"))}");
            await auditWriter.FlushAsync();

            return total;
        }

        // Report 0911 can be slow or occasionally not return a result. If the result (or total) doesn't
        // appear, click Run Report again (re-applying the date filters) up to MaxRuns times.
        private async Task<Report911RangeResult> RunReport911RangeAsync(
            DateTime startDate,
            DateTime endDate,
            bool hasEditableEndDateFilter)
        {
            const int MaxRuns = 3;
            Report911RangeResult? last = null;
            for (int run = 1; run <= MaxRuns; run++)
            {
                try
                {
                    last = await RunReport911AttemptAsync(startDate, endDate, hasEditableEndDateFilter);
                    if (last.Amount != null || last.NoRows || run == MaxRuns)
                        return last;
                    Console.WriteLine(
                        $"Report 911 {startDate:MM/dd/yyyy} - {endDate:MM/dd/yyyy}: no total (run {run}); clicking Run Report again.");
                }
                catch (Exception ex) when (run < MaxRuns && (ex is TimeoutException || ex is PlaywrightException))
                {
                    Console.WriteLine(
                        $"Report 911 {startDate:MM/dd/yyyy} - {endDate:MM/dd/yyyy}: not loaded (run {run}): {ex.Message.Split('\n')[0]}; clicking Run Report again.");
                    try { await WaitForLoadingOverlayToDisappearAsync(timeout: 15000); } catch { }
                }
            }
            return last!;
        }


        private async Task<Report911RangeResult> RunReport911AttemptAsync(
            DateTime startDate,
            DateTime endDate,
            bool hasEditableEndDateFilter)
        {
            await TransStartDateFilter.WaitForAsync(new() { State = WaitForSelectorState.Visible });
            await Btn911Filter.WaitForAsync(new() { State = WaitForSelectorState.Visible });
            await SetReportDateFilterAsync(TransStartDateFilter, startDate, "start");
            if (hasEditableEndDateFilter)
                await SetReportDateFilterAsync(TransEndDateFilter, endDate, "end");

            if (!await Btn911Filter.IsEnabledAsync())
                throw new InvalidOperationException("Report 911 Run Report button is disabled.");

            await EnsureShowStaffCheckedAsync();
            var reportWait = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                await Btn911Filter.ClickAsync(new() { NoWaitAfter = true });
                var navigationTimeout = Math.Max(
                    1,
                    MaximumUiWaitTimeoutMs - (int)reportWait.ElapsedMilliseconds);
                await _page.WaitForLoadStateAsync(
                    LoadState.DOMContentLoaded,
                    new() { Timeout = navigationTimeout });
            }
            catch (TimeoutException ex)
            {
                throw new TimeoutException(
                    $"Report 911 navigation did not reach DOMContentLoaded within " +
                    $"{MaximumUiWaitTimeoutMs}ms for {startDate:MM/dd/yyyy} - {endDate:MM/dd/yyyy}.",
                    ex);
            }

            var hasNoData = false;
            try
            {
                var remainingTimeout = Math.Max(
                    1,
                    MaximumUiWaitTimeoutMs - (int)reportWait.ElapsedMilliseconds);
                await WaitForLoadingOverlayToDisappearAsync(
                    timeout: remainingTimeout,
                    readyLocator: Table911);

                remainingTimeout = Math.Max(
                    1,
                    MaximumUiWaitTimeoutMs - (int)reportWait.ElapsedMilliseconds);
                await WaitForLocatorWithFallbackAsync(Table911, timeout: remainingTimeout);
            }
            catch (TimeoutException ex)
            {
                var pageText = await _page.Locator("body").InnerTextAsync();
                if (!IsReport911NoDataMessage(pageText))
                {
                    throw new TimeoutException(
                        $"Report 911 did not show a result table or a no-data message for " +
                        $"{startDate:MM/dd/yyyy} - {endDate:MM/dd/yyyy}.",
                        ex);
                }

                hasNoData = true;
            }

            if (!hasNoData)
            {
                var emptyResult = Table911.Locator("tbody td.dataTables_empty");
                var emptyResultCount = await emptyResult.CountAsync();
                var resultText = await Table911.InnerTextAsync();
                hasNoData = (emptyResultCount > 0 && await emptyResult.First.IsVisibleAsync())
                    || IsReport911NoDataMessage(resultText);
            }

            if (hasNoData)
                return new Report911RangeResult(0m, 0, true, null);

            var bodyRowCount = await Table911.Locator("tbody tr").CountAsync();
            var footerCount = await Report911Total.CountAsync();
            if (bodyRowCount > 0 && footerCount == 0)
            {
                // Rows are shown but the total footer is still rendering; give it time.
                try
                {
                    await Report911Total.First.WaitForAsync(
                        new() { State = WaitForSelectorState.Attached, Timeout = 20000 });
                }
                catch (TimeoutException)
                {
                }
                footerCount = await Report911Total.CountAsync();
            }
            if (bodyRowCount == 0 && footerCount == 0)
                return new Report911RangeResult(0m, 0, true, null);

            decimal rangeAmount = 0m;
            string? amountText = null;
            foreach (var footerText in await Report911Total.AllTextContentsAsync())
            {
                var candidate = footerText.Trim();
                if (TryParseReport911Amount(candidate, out rangeAmount))
                {
                    amountText = candidate;
                    break;
                }
            }

            if (amountText == null)
            {
                var footerAmounts = new HashSet<decimal>();
                foreach (var footerText in await Report911FooterCells.AllTextContentsAsync())
                {
                    if (TryParseReport911Amount(footerText.Trim(), out var footerAmount))
                        footerAmounts.Add(footerAmount);
                }

                if (footerAmounts.Count == 1)
                {
                    rangeAmount = footerAmounts.Single();
                    amountText = rangeAmount.ToString(CultureInfo.InvariantCulture);
                }
            }

            if (amountText == null)
            {
                var tableDiagnostics = await Table911.EvaluateAsync<string>(
                    @"table => JSON.stringify({
                        className: table.className,
                        headerTexts: Array.from(table.querySelectorAll('thead th'),
                            cell => cell.textContent.trim()),
                        tableRows: Array.from(table.ownerDocument.querySelectorAll('table'),
                            candidate => ({
                                id: candidate.id,
                                className: candidate.className,
                                rowCount: candidate.querySelectorAll('tbody tr').length,
                                footerCount: candidate.querySelectorAll('tfoot').length,
                                footerCellTexts: Array.from(candidate.querySelectorAll('tfoot td'),
                                    cell => cell.textContent.trim()),
                                headers: Array.from(candidate.querySelectorAll('thead th'),
                                    cell => cell.textContent.trim())
                            })),
                        bodyRowCount: table.querySelectorAll('tbody tr').length,
                        firstRowCellCount: table.querySelector('tbody tr')?.cells.length ?? 0,
                        columnSummaries: (() => {
                            const rows = Array.from(table.querySelectorAll('tbody tr'));
                            const columnCount = rows[0]?.cells.length ?? 0;
                            return Array.from({ length: columnCount }, (_, index) => {
                                const values = rows.map(row =>
                                    (row.cells[index]?.textContent ?? '').trim());
                                const currencyValues = values.filter(value =>
                                    /[$,.]/.test(value)
                                    && /^\(?-?\$?\s*\d[\d,]*(?:\.\d{1,2})?\)?$/.test(value));
                                const sum = currencyValues.reduce((total, value) => {
                                    const normalized = value.replace(/[$,\s]/g, '');
                                    const amount = normalized.startsWith('(')
                                        ? -Number(normalized.slice(1, -1))
                                        : Number(normalized);
                                    return total + amount;
                                }, 0);
                                return {
                                    index,
                                    currencyValueCount: currencyValues.length,
                                    currencySum: sum
                                };
                            });
                        })(),
                        footerCount: table.querySelectorAll('tfoot').length,
                        footerCellCount: table.querySelectorAll('tfoot td').length,
                        footerCellTexts: Array.from(table.querySelectorAll('tfoot td'),
                            cell => cell.textContent.trim()),
                        totalLabelCount: Array.from(table.querySelectorAll('td, th'))
                            .filter(cell => /^\s*total\s*:?\s*$/i.test(cell.textContent)).length,
                        dataTablesApiInfo: (() => {
                            const jq = table.ownerDocument.defaultView?.jQuery;
                            return jq?.fn?.dataTable?.isDataTable(table)
                                ? jq(table).DataTable().page.info()
                                : null;
                        })()
                    })");
                return new Report911RangeResult(
                    null,
                    bodyRowCount,
                    false,
                    tableDiagnostics);
            }

            return new Report911RangeResult(rangeAmount, bodyRowCount, false, null);
        }

        private static bool TryParseReport911Amount(string text, out decimal amount)
        {
            var culture = CultureInfo.GetCultureInfo("en-US");
            if (decimal.TryParse(text, NumberStyles.Currency, culture, out amount))
                return true;

            var match = System.Text.RegularExpressions.Regex.Match(
                text,
                @"\(?-?\$?\s*(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}\)?");
            return match.Success
                && decimal.TryParse(match.Value, NumberStyles.Currency, culture, out amount);
        }

        private static bool IsDateValue(DateTime expectedDate, string actualValue)
        {
            return DateTime.TryParse(
                       actualValue,
                       CultureInfo.InvariantCulture,
                       DateTimeStyles.None,
                       out var parsedDate)
                   && parsedDate.Date == expectedDate.Date;
        }



        // This method navigates to the Fee Management section
        public async Task NavigateToFeeManagementAsync()
        {
            // Wait after dismissing optional popups, which can keep the loader active.
            await WaitForLoadingOverlayToDisappearAsync();

            // The Administration dropdown sometimes does not open (or closes again) on a slow page,
            // leaving Fee Management in the DOM but hidden. Retry opening it, then fall back to its link.
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                await AdministrationMenu.First.ClickAsync();
                try
                {
                    await FeeManagementMenu.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 8000 });
                    await FeeManagementMenu.ClickAsync(new() { Timeout = 10000 });
                    return;
                }
                catch (TimeoutException)
                {
                    Console.WriteLine($"Fee Management menu not clickable (attempt {attempt}); retrying.");
                    await WaitForLoadingOverlayToDisappearAsync();
                }
            }

            string href = await FeeManagementMenu.GetAttributeAsync("href");
            if (string.IsNullOrWhiteSpace(href))
            {
                throw new InvalidOperationException("Fee Management menu link could not be opened.");
            }
            await _page.GotoAsync(new Uri(new Uri(_page.Url), href).ToString());
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
            await FeeReportsSearch.WaitForAsync(new() { State = WaitForSelectorState.Visible });
        }


        // This method verifies whether the Fee Management dashboard is displayed
        public async Task<(bool, long)> SearchFeeReportAsync(long reportId, IPage page)
        {
            // Wait for the search box
            await FeeReportsSearch.WaitForAsync(new LocatorWaitForOptions
            {
                State = WaitForSelectorState.Visible,
                Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
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
            // The report opens in a new tab; await its DOM instead of sleeping for a fixed duration.
            var reportPageTask = _page.Context.WaitForPageAsync();
            await FeeReportsFirst.ClickAsync();
            var reportPage = await reportPageTask;
            PlaywrightTimeoutPolicy.Configure(reportPage.Context);
            await reportPage.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
            await MaxWindow();
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
                await _page.EvaluateAsync(
                    "window.scrollBy(0, document.documentElement ? document.documentElement.scrollHeight : 0)");
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

        //#region Verify Account Status Report data #907
        //public async Task<Report> VerifyAccountStatusAsync(Report objreport)
        //{
        //    try
        //    {
        //        await MaxWindow();
        //        await _page.GotoAsync("https://feemanagement.optionc.com/report-list", new() { Timeout = 100000 });
        //        await SearchFeeReportAsync(907, _page);
        //        await MaxWindow();
        //        var pages = _page.Context.Pages;
        //        // Ensure multiple pages exist before switching
        //        if (pages.Count > 1)
        //        {
        //            var firstPage = pages.FirstOrDefault();
        //            if (firstPage != null && !firstPage.IsClosed)
        //            {
        //                await firstPage.CloseAsync();
        //            }
        //            // Switch to the last opened page and update _page
        //            _page = _page.Context.Pages.LastOrDefault() ?? throw new Exception("Failed to switch to the last opened page.");
        //            await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        //        }
        //        // Click on StaffCheck if visible
        //        if (await StaffCheck.IsVisibleAsync())
        //        {
        //            // await StaffCheck.ClickAsync(new() { Timeout = 10000 });
        //            // await _page.WaitForTimeoutAsync(30000);

        //            await StaffCheck.ClickAsync();
        //            await _page.WaitForSelectorAsync("//input[@value='Run Report']", new() { State = WaitForSelectorState.Visible, Timeout = 30000 });

        //        }
        //        // Ensure 'Run Report' button is visible before clicking
        //        // await _page.WaitForSelectorAsync("//input[@value='Run Report']", new() { State = WaitForSelectorState.Visible });
        //        //if (await _page.Locator("//input[@value='Run Report']").IsVisibleAsync())
        //        //{
        //        //    // Go ahead and click
        //        //    await BtnRunReport.ClickAsync();
        //        //}

        //        //var runReportLocator = _page.Locator("//input[@value='Run Report']");
        //        var runReportLocator = _page.Locator("//input[@type='button' and @value='Run Report' and contains(@class,'btn-info')]");
        //        await runReportLocator.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout =50000 });

        //        if (await runReportLocator.IsVisibleAsync())
        //        {
        //            await BtnRunReport.ClickAsync();
        //        }


        //        //if (await BtnRunReport.IsEnabledAsync())
        //        //{
        //        //    await BtnRunReport.ClickAsync(new() { Timeout = 100000 });
        //        //    await _page.WaitForTimeoutAsync(100000);
        //        //    Console.WriteLine("Run Report button clicked.");
        //        //}
        //        else
        //        {
        //            throw new Exception("Run Report button is disabled or not interactable.");
        //        }
        //        // Wait for dropdown visibility
        //        await _page.WaitForSelectorAsync("//*[@class='form-control input-xsmall input-inline']", new() { State = WaitForSelectorState.Visible });

        //        if (await RecordLength.IsVisibleAsync())
        //        {
        //            await RecordLength.SelectOptionAsync("-1");
        //            Console.WriteLine("Record Length dropdown set to -1.");
        //        }
        //        else
        //        {
        //            throw new Exception("Record Length dropdown not found.");
        //        }
        //        await _page.WaitForTimeoutAsync(2000); // Allow table update
        //        // Fetch row counts
        //        var primaryRows = await _page.Locator("//table//tbody//tr").CountAsync();
        //        var rowsCount = await _page.Locator("//*[@id='tblParentAccountSetup']/tbody/tr/td[@class='dataTables_empty']").CountAsync();
        //        objreport.TotalPrimaryAccount = (rowsCount != 1) ? primaryRows.ToString() : "0";
        //        await using (StreamWriter writer = new StreamWriter(path, true))
        //        {
        //            await writer.WriteLineAsync($"\r\nACCOUNT SETUP STATISTICS\r\n========================\r\nTotal Primary Accounts: {objreport.TotalPrimaryAccount}");
        //        }
        //        // Ensure dropdown selection again
        //        if (await RecordLength.IsVisibleAsync())
        //        {
        //            await RecordLength.SelectOptionAsync("-1");
        //        }
        //        // Click 'Account Created' radio button if visible
        //        if (await AccountCreatedRadio.IsVisibleAsync())
        //        {
        //            await AccountCreatedRadio.ClickAsync();
        //            Console.WriteLine("Account Created Radio button clicked.");
        //        }
        //        // Click 'Run Report' again
        //        if (await BtnRunReport.IsEnabledAsync())
        //        {
        //            await BtnRunReport.ClickAsync(new() { Timeout = 10000 });
        //            Console.WriteLine("Run Report clicked again.");
        //        }
        //        else
        //        {
        //            Console.WriteLine("Run Report button is disabled.");
        //        }
        //        await _page.WaitForTimeoutAsync(2000); // Allow table update
        //        // Fetch updated row count
        //        var totalAccountsCreated = await _page.Locator("//table//tbody//tr").CountAsync();
        //        var rowsCount1 = await _page.Locator("//*[@id='tblParentAccountSetup']/tbody/tr/td[@class='dataTables_empty']").CountAsync();
        //        objreport.TotalAccountCreated = (rowsCount1 != 1) ? totalAccountsCreated.ToString() : "0";
        //        await using (StreamWriter writer = new StreamWriter(path, true))
        //        {
        //            await writer.WriteLineAsync($"Total Accounts Created: {objreport.TotalAccountCreated}");
        //        }
        //        // Extracting yesterday’s date
        //        string yesterdayString = DateTime.Now.AddDays(-1).ToString("MM/dd/yyyy");
        //        long creditCardCount = 0, eCheckCount = 0;
        //        if (rowsCount1 != 1)
        //        {
        //            var rows = await _page.Locator("//table//tbody//tr").AllAsync();
        //            foreach (var row in rows)
        //            {
        //                var status = await row.Locator("td:nth-of-type(7)").InnerTextAsync();
        //                var dateText = await row.Locator("td:nth-of-type(8)").InnerTextAsync();
        //                var typeText = await row.Locator("td:nth-of-type(6)").InnerTextAsync();
        //                var nameText = await row.Locator("td:nth-of-type(4)").InnerTextAsync();

        //                if (DateTime.TryParse(dateText, out DateTime dateParsed))
        //                {
        //                    string dateFormatted = dateParsed.ToString("MM/dd/yyyy");

        //                    if ((status.Contains("In Progress")) || dateFormatted == yesterdayString)
        //                    {
        //                        string accountStatus = $"Name: {nameText} ; Date: {dateText} ; Status: {status} ; Account Type: {typeText} ;";

        //                        await using (StreamWriter writer = new StreamWriter(path, true))
        //                        {
        //                            await writer.WriteLineAsync(accountStatus);
        //                        }

        //                        objreport.AccountStatus += accountStatus + "\n";
        //                    }

        //                    if (typeText == "Credit Card") creditCardCount++;
        //                    else eCheckCount++;
        //                }
        //            }
        //        }
        //        objreport.CreditCardCount = creditCardCount.ToString();
        //        objreport.ECheckCount = eCheckCount.ToString();
        //        await using (StreamWriter writer = new StreamWriter(path, true))
        //        {
        //            await writer.WriteLineAsync($"Total CC Accounts Created: {creditCardCount}\r\nTotal eCheck Accounts Created: {eCheckCount}");
        //            await writer.WriteLineAsync("###############################################################################");
        //        }
        //        return objreport;
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error encountered: {ex.Message}");
        //        throw;
        //    }
        //}
        //#endregion


        #region Account Activated On , Verify Account Status Report data #907

        // Opens Account History in Report 0907 and counts accounts activated in the last 24 hours (school time).
        // Expects the Report 0907 page to be the current page, already run once with the Show boxes ticked.
        private async Task<int> CountActivatedLast24HoursAsync()
        {
                LogToFile("Checking Account History - Accounts Activated in Last 24 Hours", true);

                // Select Account History radio button
                await RbAccountHistory.WaitForAsync(
                    new() { State = WaitForSelectorState.Visible });
                await RbAccountHistory.ClickAsync();
                LogToFile("Account History radio button selected.");

                // Run Account History report
                await EnsureShowStaffCheckedAsync();
                if (await BtnRunReport.IsEnabledAsync())
                {
                    await BtnRunReport.ClickAsync(new()
                    {
                        Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
                    });

                    LogToFile("Run Report clicked for Account History.");
                    await WaitForLoadingOverlayToDisappearAsync();
                }
                else
                {
                    throw new Exception("Run Report button is disabled for Account History.");
                }

                // Wait for Account History grid
                await ThActivatedOn.WaitForAsync(
                    new()
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
                    });

                // Display all Account History records
                if (await RecordLength.IsVisibleAsync())
                {
                    await RecordLength.SelectOptionAsync("-1");

                    LogToFile("Account History Record Length set to All.");
                }
                await WaitForLoadingOverlayToDisappearAsync();

                // Sort Activated On column in descending order
                var sortOrder = await ThActivatedOn.GetAttributeAsync("aria-sort");

                if (sortOrder != "descending")
                {
                    await ThActivatedOn.ClickAsync();

                    sortOrder = await ThActivatedOn.GetAttributeAsync("aria-sort");

                    // If first click sorts ascending, click again for descending
                    if (sortOrder != "descending")
                    {
                        await ThActivatedOn.ClickAsync();
                    }
                }

                LogToFile("Activated On column sorted in descending order.");

                // Calculate the previous 24 hours
                // The "Activated On" times are US (school) times, while this machine runs on India time.
                // Compare them against the current time in the US zone, not the machine's local time.
                DateTime currentTime = GetSchoolLocalNow();

                DateTime last24Hours = currentTime.AddHours(-24);
                LogToFile($"Counting activations between {last24Hours:MM/dd/yyyy hh:mm:ss tt} and {currentTime:MM/dd/yyyy hh:mm:ss tt} (school time).");

                int activatedCount = 0;

                // Get all Account History rows
                var accountHistoryRows = _page.Locator(
                    "//table[@id='tblParentAccountSetup']/tbody/tr");

                int historyRowCount = await accountHistoryRows.CountAsync();

                LogToFile($"Total Account History records: {historyRowCount}");

                // Read Activated On date from the 9th column
                for (int i = 0; i < historyRowCount; i++)
                {
                    var row = accountHistoryRows.Nth(i);

                    // Skip empty table message
                    if (await row.Locator("td.dataTables_empty").CountAsync() > 0)
                    {
                        continue;
                    }

                    var activatedCell = row.Locator("td").Nth(8);

                    // Read data-sort attribute
                    string activatedDateText =
                        await activatedCell.GetAttributeAsync("data-sort") ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(activatedDateText))
                    {
                        continue;
                    }

                    // Parse date format: 20260817 05:00:00 AM
                    if (DateTime.TryParseExact(
                        activatedDateText.Trim(),
                        "yyyyMMdd hh:mm:ss tt",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out DateTime activatedDate))
                    {
                        // Check whether activation happened within the last 24 hours
                        if (activatedDate >= last24Hours &&
                            activatedDate <= currentTime)
                        {
                            activatedCount++;

                            LogToFile(
                                $"Activated Account: {activatedDate:MM/dd/yyyy hh:mm:ss tt}");
                        }
                    }
                    else
                    {
                        LogToFile(
                            $"Unable to parse Activated On date: {activatedDateText}");
                    }
                }
                LogToFile(
                    $"Total Accounts Activated in the Last 24 Hours: {activatedCount}");
                return activatedCount;
        }

        public async Task<Report> VerifyAccountStatusAsync(Report objreport)
        {
            try
            {
                await MaxWindow();

                await _page.GotoAsync(
                    "https://feemanagement.optionc.com/report-list",
                    new() { Timeout = PlaywrightTimeoutPolicy.NavigationTimeoutMs });

                await SearchFeeReportAsync(907, _page);

                await MaxWindow();

                var pages = _page.Context.Pages;

                // Switch to the newly opened report page
                if (pages.Count > 1)
                {
                    var firstPage = pages.FirstOrDefault();

                    if (firstPage != null && !firstPage.IsClosed)
                    {
                        await firstPage.CloseAsync();
                    }

                    _page = _page.Context.Pages.LastOrDefault()
                        ?? throw new Exception("Failed to switch to the last opened page.");

                    PlaywrightTimeoutPolicy.Configure(_page.Context);
                    await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                }

                // Run Report
                var runReportLocator = _page.Locator(
                    "//input[@type='button' and @value='Run Report' and contains(@class,'btn-info')]");

                await runReportLocator.WaitForAsync(
                    new()
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
                    });

                if (await runReportLocator.IsVisibleAsync() &&
                    await runReportLocator.IsEnabledAsync())
                {
                    await EnsureShowStaffCheckedAsync();
                    await BtnRunReport.ClickAsync();
                    await WaitForLoadingOverlayToDisappearAsync();
                }
                else
                {
                    throw new Exception("Run Report button is disabled or not interactable.");
                }

                // Wait for Record Length dropdown
                await RecordLength.WaitForAsync(
                    new()
                    {
                        State = WaitForSelectorState.Visible,
                        Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
                    });

                // Display all records
                await RecordLength.SelectOptionAsync("-1");
                await WaitForLoadingOverlayToDisappearAsync();

                // --------------------------------------------------
                // 1. TOTAL PRIMARY ACCOUNTS
                // --------------------------------------------------

                var primaryRows = await GetParentAccountRowCountAsync();
                objreport.TotalPrimaryAccount = primaryRows.ToString();

                await using (StreamWriter writer = new StreamWriter(path, true))
                {
                    await writer.WriteLineAsync(
                        $"\r\nACCOUNT SETUP STATISTICS\r\n" +
                        $"========================\r\n" +
                        $"Total Primary Accounts: {objreport.TotalPrimaryAccount}");
                }

                // --------------------------------------------------
                // 2. ACCOUNT CREATED REPORT
                // --------------------------------------------------

                await RunAccountCreatedReportAsync();

                // Fetch Account Created row count
                var totalAccountsCreated = await GetParentAccountRowCountAsync();
                objreport.TotalAccountCreated = totalAccountsCreated.ToString();

                await using (StreamWriter writer = new StreamWriter(path, true))
                {
                    await writer.WriteLineAsync(
                        $"Total Accounts Created: {objreport.TotalAccountCreated}");
                }

                // --------------------------------------------------
                // 3. ACCOUNT CREATED DETAILS AND PAYMENT TYPE COUNT
                // --------------------------------------------------

                long creditCardCount = 0;
                long eCheckCount = 0;

                if (totalAccountsCreated > 0)
                {
                    var rows = await _page.Locator(
                        "//table[@id='tblParentAccountSetup']/tbody/tr").AllAsync();

                    foreach (var row in rows)
                    {
                        var typeText = (await row.Locator("td:nth-of-type(6)").InnerTextAsync()).Trim();

                        if (typeText.Equals("Credit Card", StringComparison.OrdinalIgnoreCase))
                        {
                            creditCardCount++;
                        }
                        else if (typeText.Equals("eCheck", StringComparison.OrdinalIgnoreCase))
                        {
                            eCheckCount++;
                        }
                    }
                }

                objreport.CreditCardCount = creditCardCount.ToString();
                objreport.ECheckCount = eCheckCount.ToString();

                await using (StreamWriter writer = new StreamWriter(path, true))
                {
                    await writer.WriteLineAsync(
                        $"Total CC Accounts Created: {creditCardCount}\r\n" +
                        $"Total eCheck Accounts Created: {eCheckCount}");
                }

                await CollectNewOrInProgressAccountsAsync(objreport);

                // --------------------------------------------------
                // 4. ACCOUNT HISTORY - ACTIVATED IN LAST 24 HOURS
                // --------------------------------------------------

                int activatedCount = await CountActivatedLast24HoursAsync();
                objreport.AcctsActivatedLast24Hours = activatedCount.ToString();

                // Write result to log file
                await using (StreamWriter writer = new StreamWriter(path, true))
                {
                    await writer.WriteLineAsync(
                        $"Accts Activated (Last 24 hrs): {objreport.AcctsActivatedLast24Hours}");

                    await writer.WriteLineAsync(
                        "###############################################################################");
                }

                return objreport;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error encountered: {ex.Message}");

                LogToFile($"Error in Account Status Report: {ex.Message}");

                throw;
            }
        }

        private static bool IsNewOrInProgressStatus(string status)
        {
            var normalizedStatus = string.Concat(status.Where(char.IsLetterOrDigit));
            return normalizedStatus.StartsWith("New", StringComparison.OrdinalIgnoreCase)
                || normalizedStatus.StartsWith("InProgress", StringComparison.OrdinalIgnoreCase);
        }

        private async Task CollectNewOrInProgressAccountsAsync(Report report)
        {
            if (!await AccountCreatedRadio.IsCheckedAsync())
            {
                await RunAccountCreatedReportAsync();
            }

            var table = _page.Locator("#tblParentAccountSetup");
            await table.WaitForAsync(
                new()
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
                });

            var headerTexts = await table.Locator("thead th").AllInnerTextsAsync();
            var statusColumn = FindAccountColumn(headerTexts, "status", 6);
            if (statusColumn < 0)
            {
                throw new InvalidOperationException(
                    $"The Account Created report does not expose a sortable Status column. " +
                    $"Available columns: {string.Join(" | ", headerTexts)}");
            }

            var statusHeader = table.Locator("thead th").Nth(statusColumn);
            var sortOrder = await statusHeader.GetAttributeAsync("aria-sort");
            if (!string.Equals(sortOrder, "ascending", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(sortOrder, "descending", StringComparison.OrdinalIgnoreCase))
            {
                await statusHeader.ClickAsync();
                await WaitForLoadingOverlayToDisappearAsync();
                sortOrder = await statusHeader.GetAttributeAsync("aria-sort");
            }

            if (!string.Equals(sortOrder, "ascending", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(sortOrder, "descending", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "The Account Created report's Status column could not be sorted.");
            }

            var nameColumn = FindAccountColumn(headerTexts, "name", 3);
            var typeColumn = FindAccountColumn(headerTexts, "type", 5);
            var dateColumn = FindAccountColumn(headerTexts, "date", 7);
            var rows = table.Locator("tbody tr");
            var rowCount = await rows.CountAsync();
            var matchingAccountCount = 0;

            for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var row = rows.Nth(rowIndex);
                if (await row.Locator("td.dataTables_empty").CountAsync() > 0)
                {
                    continue;
                }

                var cells = row.Locator("td");
                var status = (await cells.Nth(statusColumn).InnerTextAsync()).Trim();
                if (!IsNewOrInProgressStatus(status))
                {
                    continue;
                }

                var name = (await cells.Nth(nameColumn).InnerTextAsync()).Trim();
                var date = (await cells.Nth(dateColumn).InnerTextAsync()).Trim();
                var accountType = (await cells.Nth(typeColumn).InnerTextAsync()).Trim();
                var accountStatus =
                    $"Name: {name} ; Date: {date} ; Status: {status} ; Account Type: {accountType} ;";

                await File.AppendAllTextAsync(path, accountStatus + Environment.NewLine);
                report.AccountStatus += accountStatus + Environment.NewLine;
                matchingAccountCount++;
            }

            Console.WriteLine(
                $"New/InProgress accounts found in Account Created report: {matchingAccountCount}");
        }

        private async Task RunAccountCreatedReportAsync()
        {
            await AccountCreatedRadio.WaitForAsync(
                new()
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
                });

            if (!await AccountCreatedRadio.IsCheckedAsync())
            {
                await AccountCreatedRadio.ClickAsync();
            }

            await EnsureShowStaffCheckedAsync();
            if (!await BtnRunReport.IsEnabledAsync())
            {
                throw new InvalidOperationException(
                    "Run Report is disabled for the Account Created report.");
            }

            await BtnRunReport.ClickAsync(
                new() { Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });
            await WaitForLoadingOverlayToDisappearAsync();

            if (await RecordLength.IsVisibleAsync())
            {
                await RecordLength.SelectOptionAsync("-1");
                await WaitForLoadingOverlayToDisappearAsync();
            }
        }

        private static int FindAccountColumn(
            IReadOnlyList<string> headers,
            string headerToken,
            int fallbackIndex)
        {
            for (var index = 0; index < headers.Count; index++)
            {
                if (headers[index].Contains(headerToken, StringComparison.OrdinalIgnoreCase))
                {
                    return index;
                }
            }

            if (fallbackIndex >= 0 && fallbackIndex < headers.Count)
            {
                return fallbackIndex;
            }

            return -1;
        }

        private async Task<int> GetParentAccountRowCountAsync()
        {
            var table = _page.Locator("#tblParentAccountSetup");
            await table.WaitForAsync(
                new()
                {
                    State = WaitForSelectorState.Visible,
                    Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
                });

            return await table.EvaluateAsync<int>(
                @"table => {
                    const jquery = window.jQuery;
                    if (jquery && jquery.fn && jquery.fn.dataTable &&
                        jquery.fn.dataTable.isDataTable(table)) {
                        return jquery(table).DataTable().page.info().recordsDisplay;
                    }

                    return Array.from(table.querySelectorAll('tbody tr'))
                        .filter(row => !row.querySelector('td.dataTables_empty'))
                        .length;
                }");
        }

        // ---------------- TEMPORARY: refresh only "Acts Activated (Last 24 hrs)" ----------------
        private bool _suppressTextLog;

        /* Temporary logic - commented out; uncomment (and the step in MMAuditStepDefinitions) to re-run it.
        // Opens Report 0907 for the current school, counts activations (corrected logic) and writes ONLY
        // column S of that school's existing row. Nothing else in the workbook or the text log is touched.
        public async Task RefreshActivatedLast24HoursAsync(string schoolCode)
        {
            _suppressTextLog = true;
            try
            {
                await MaxWindow();
                await _page.GotoAsync(
                    "https://feemanagement.optionc.com/report-list",
                    new() { Timeout = PlaywrightTimeoutPolicy.NavigationTimeoutMs });
                await SearchFeeReportAsync(907, _page);
                await MaxWindow();

                if (_page.Context.Pages.Count > 1)
                {
                    var firstPage = _page.Context.Pages.FirstOrDefault();
                    if (firstPage != null && !firstPage.IsClosed)
                        await firstPage.CloseAsync();
                    _page = _page.Context.Pages.LastOrDefault()
                        ?? throw new Exception("Failed to switch to the last opened page.");
                    PlaywrightTimeoutPolicy.Configure(_page.Context);
                    await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                }

                var runReport = _page.Locator(
                    "//input[@type='button' and @value='Run Report' and contains(@class,'btn-info')]");
                await runReport.WaitForAsync(
                    new() { State = WaitForSelectorState.Visible, Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });
                await EnsureShowStaffCheckedAsync();
                await BtnRunReport.ClickAsync();
                await WaitForLoadingOverlayToDisappearAsync();

                int count = await CountActivatedLast24HoursAsync();
                UpdateActivatedColumnInExcel(schoolCode, count);
                Console.WriteLine($"Acts Activated refreshed for {schoolCode}: {count}");
            }
            finally
            {
                _suppressTextLog = false;
            }
        }

        private void UpdateActivatedColumnInExcel(string schoolCode, int count)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(auditWorkbookPath)))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.FirstOrDefault(w => w.Name == "2026-2027")
                    ?? throw new InvalidOperationException("The audit workbook has no '2026-2027' sheet.");
                int end = ws.Dimension?.End.Row ?? 1;
                int row = -1;
                for (int r = 2; r <= end; r++)
                {
                    if (string.Equals(AuditCodeAt(ws, r), schoolCode.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        row = r;
                        break;
                    }
                }
                if (row < 0)
                    throw new InvalidOperationException($"School {schoolCode} has no row in the audit workbook.");

                ws.Cells[row, 19].Value = count;
                ws.Cells[row, 19].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 19].Style.Fill.BackgroundColor.SetColor(count >= 1 ? Color.Yellow : Color.MintCream);
                package.Save();
            }
        }
        */

        // Current time in the US zone the school reports use (Central Time).
        private static DateTime GetSchoolLocalNow()
        {
            foreach (var id in new[] { "Central Standard Time", "America/Chicago" })
            {
                try
                {
                    return TimeZoneInfo.ConvertTime(DateTime.UtcNow, TimeZoneInfo.FindSystemTimeZoneById(id));
                }
                catch (TimeZoneNotFoundException) { }
                catch (InvalidTimeZoneException) { }
            }
            // Fallback: Central Standard Time (UTC-6), ignoring daylight saving.
            return DateTime.UtcNow.AddHours(-6);
        }

        private void LogToFile(string message)
        {
            if (_suppressTextLog) { Console.WriteLine(message); return; }
            try
            {
                using (StreamWriter writer = new StreamWriter(path, true))
                {
                    writer.WriteLine($"{DateTime.Now:MM/dd/yyyy hh:mm:ss tt} - {message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to log file: {ex.Message}");
            }
        }

        private void LogToFile(string message, bool addHeader)
        {
            if (_suppressTextLog) { Console.WriteLine(message); return; }
            try
            {
                using (StreamWriter writer = new StreamWriter(path, true))
                {
                    if (addHeader)
                    {
                        writer.WriteLine();
                        writer.WriteLine("==================================================");
                        writer.WriteLine($"{DateTime.Now:MM/dd/yyyy hh:mm:ss tt} - {message}");
                        writer.WriteLine("==================================================");
                    }
                    else
                    {
                        writer.WriteLine($"{DateTime.Now:MM/dd/yyyy hh:mm:ss tt} - {message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing to log file: {ex.Message}");
            }
        }
        #endregion


        #region Report Button Filter, write the report value and compare results are matching or not
        public async Task<ReportData> StartFilterAsync(
            Int64 ReportId,
            Report objReport,
            StreamWriter auditWriter,
            string schoolIdentifier)
        {

            var reportData = new ReportData();
            string getTransactionAmt = "0.00";
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
            PlaywrightTimeoutPolicy.Configure(_page.Context);
            await _page.BringToFrontAsync(); // Bring the last page to the front
            await MaxWindow();

            if (ReportId == 911)
            {
                var report911Total = await GetReport911TotalAsync(auditWriter, schoolIdentifier);
                reportData.ReportValue = report911Total.ToString("F2", CultureInfo.CurrentCulture);
                return reportData;
            }

            // Start date filtering
            //await TransStartDateFilter.ClickAsync();
            //await TransStartDateFilter.FillAsync(DateFunction("01/01/2020")); //Day 1 mm released date
            // Click inside the Start Date field
            await TransStartDateFilter.ClickAsync();

            // Clear any existing value
            await TransStartDateFilter.PressAsync("Control+A"); // Select all
            await TransStartDateFilter.PressAsync("Backspace"); // Delete existing text

            // Type the required date manually
            await TransStartDateFilter.TypeAsync(text: "01/01/2020");

            // Trigger blur/change event by tabbing out
            await TransStartDateFilter.PressAsync("Tab");




            //await WaitForElementAsync("//input[@onclick='ShowReportDetails(this.id)']");
            await EnsureShowStaffCheckedAsync();
            await BtnFilter.ClickAsync(new()
            {
                Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
            });
            await WaitForLoadingOverlayToDisappearAsync(readyLocator: Table);

            // Wait for the table to load
            await WaitForLocatorWithFallbackAsync(Table);
            IReadOnlyList<IElementHandle> tableRows = await _page.QuerySelectorAllAsync("//tfoot");
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
                totalValue = tableRows.Count > 0 ? await ReportValue.First.InnerTextAsync() : "0.00";



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
        public async Task<Report> LoopFilterAsync(string schoolIdentifier)
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
                        await _page.GotoAsync(
                            "https://feemanagement.optionc.com/report-list",
                            new() { Timeout = PlaywrightTimeoutPolicy.NavigationTimeoutMs });

                        await SearchFeeReportAsync(reportId, _page);
                        await MaxWindow();
                        ReportData reportData = await StartFilterAsync(
                            reportId,
                            objReport,
                            writer,
                            schoolIdentifier);

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
        string schoolCode = "",
        string startDate = "")
        {
            await SchoolName.WaitForAsync(); // Wait for the school name element to be present
            string getSchoolName = (await SchoolName.TextContentAsync())?.Trim() ?? "Unknown School";
            string filePath = auditWorkbookPath;
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "2026-2027")
                                              ?? package.Workbook.Worksheets.Add("2026-2027");
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
                workSheet.Columns[19].Width = 10;
                workSheet.Columns[20].Width = 63;
                workSheet.Columns[21].Width = 65;
                EnsureAuditHeaders(workSheet);
                // The school list varies per run, so the row is found/added by school code (never by list position).
                ClearAuditTotals(workSheet);
                int lastRow = GetOrAddAuditRow(workSheet, schoolCode);
                workSheet.Cells[lastRow, 1, lastRow, 21].Clear();
                // Assigning values
                workSheet.Cells[$"A{lastRow}"].Value = lastRow - 1;
                // Show the school code with the name, e.g. "Holy Trinity School - Kewaskum (7127)".
                string codeSuffix = $"({schoolCode})";
                workSheet.Cells[$"B{lastRow}"].Value =
                    string.IsNullOrWhiteSpace(schoolCode)
                    || getSchoolName.EndsWith(codeSuffix, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(getSchoolName, schoolCode, StringComparison.OrdinalIgnoreCase)
                        ? getSchoolName
                        : $"{getSchoolName} {codeSuffix}";
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
                workSheet.Cells[$"S{lastRow}"].Value = ConvertToInt(objReport?.AcctsActivatedLast24Hours ?? "0");
                // One entry per line, no blank lines and no trailing line break after the last entry.
                workSheet.Cells[$"T{lastRow}"].Value = NormalizeMultilineCell(objReport?.AccountStatus);
                workSheet.Cells[$"U{lastRow}"].Value = NormalizeMultilineCell(objReport?.CompareResults);
                workSheet.Cells[$"T{lastRow}"].Style.WrapText = true;
                workSheet.Cells[$"U{lastRow}"].Style.WrapText = true;
                workSheet.Columns[21].Width = 105; // wide enough for each "Funding Amount ... Transaction Amount ..." line to stay on one line
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
                // Funding Amount vs Transaction Amount (902 & 905): red font for any "NOT equal" line.
                ApplyCompareResultColors(workSheet.Cells[$"U{lastRow}"]);
                // Summary Rows (placed right after the last school row)
                int finalRow = WriteAuditTotals(workSheet);
                // Formatting
                workSheet.Cells[$"A1:T1"].Style.WrapText = true;
                workSheet.Cells[$"R2:U{finalRow}"].Style.WrapText = true;
                workSheet.Cells[$"A1:T{finalRow}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                workSheet.Cells[$"A1:T{finalRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                workSheet.View.ShowGridLines = false;
                // Applying Borders
                using (ExcelRange range = workSheet.Cells[$"A1:U{finalRow}"])
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

        public async Task RecordSchoolAuditFailureAsync(
            string schoolIdentifier,
            string startDate,
            Exception exception)
        {
            string errorMessage = string.Join(
                " ",
                (exception.Message ?? exception.GetType().Name)
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
            await File.AppendAllTextAsync(
                path,
                $"AUDIT FAILED - School: {schoolIdentifier}; Start date: {startDate}; Error: {errorMessage}{Environment.NewLine}");

            using (ExcelPackage package = new ExcelPackage(new FileInfo(auditWorkbookPath)))
            {
                ExcelWorksheet workSheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "2026-2027")
                                              ?? package.Workbook.Worksheets.Add("2026-2027");
                EnsureAuditHeaders(workSheet);
                ClearAuditTotals(workSheet);

                int row = GetOrAddAuditRow(workSheet, schoolIdentifier);
                workSheet.Cells[row, 1, row, 21].Clear();
                workSheet.Cells[$"A{row}"].Value = row - 1;
                workSheet.Cells[$"B{row}"].Value = schoolIdentifier;
                workSheet.Cells[$"C{row}"].Value = startDate;
                workSheet.Cells[$"U{row}"].Value = $"FAILED: {errorMessage}";
                workSheet.Cells[$"U{row}"].Style.Font.Color.SetColor(Color.Red);

                WriteAuditTotals(workSheet);
                package.Save();
            }
        }

        // Trims each line, drops blank lines and joins with a single line break (no trailing blank line).
        // Lines containing "NOT equal" are shown in red; "are equal" lines stay black.
        private static void ApplyCompareResultColors(ExcelRange cell)
        {
            string text = cell.Text ?? cell.Value?.ToString() ?? string.Empty;
            text = cell.Value?.ToString() ?? text;
            if (text.IndexOf("not equal", StringComparison.OrdinalIgnoreCase) < 0)
                return;

            var lines = text.Split('\n');
            cell.RichText.Clear();
            cell.IsRichText = true;
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].TrimEnd('\r');
                var part = cell.RichText.Add(i < lines.Length - 1 ? line + "\n" : line);
                part.Size = 11;
                part.Bold = false;
                part.Color = line.IndexOf("not equal", StringComparison.OrdinalIgnoreCase) >= 0
                    ? Color.Red
                    : Color.Black;
            }
        }

        private static string NormalizeMultilineCell(string? text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            return string.Join("\n",
                text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(line => line.Trim())
                    .Where(line => line.Length > 0));
        }

        // ---- Row helpers: rows are keyed by the school code shown at the end of School Name, e.g. "Name (7127)",
        // so list size/order can vary per run. A failed row holds just the identifier in School Name. ----
        private const int AuditLastColumn = 21; // U

        private static string AuditCodeAt(ExcelWorksheet ws, int row)
        {
            // Legacy workbooks (created earlier today) still have a "School Code" column V; honour it until migrated.
            if (string.Equals((ws.Cells[1, 22].Text ?? string.Empty).Trim(), "School Code", StringComparison.OrdinalIgnoreCase))
            {
                string legacy = (ws.Cells[row, 22].Text ?? string.Empty).Trim();
                if (legacy.Length > 0) return legacy;
            }

            string name = (ws.Cells[row, 2].Text ?? string.Empty).Trim();
            if (name.Length == 0) return string.Empty;
            var match = System.Text.RegularExpressions.Regex.Match(name, @"\(([^()]*)\)\s*$");
            return match.Success && match.Groups[1].Value.Trim().Length > 0
                ? match.Groups[1].Value.Trim()
                : name;
        }

        // Moves the code from the old "School Code" column V into School Name, then removes column V.
        private static void MigrateLegacyCodeColumn(ExcelWorksheet ws)
        {
            if (!string.Equals((ws.Cells[1, 22].Text ?? string.Empty).Trim(), "School Code", StringComparison.OrdinalIgnoreCase))
                return;

            int end = ws.Dimension?.End.Row ?? 1;
            for (int r = 2; r <= end; r++)
            {
                string code = (ws.Cells[r, 22].Text ?? string.Empty).Trim();
                if (code.Length == 0) continue;
                string name = (ws.Cells[r, 2].Text ?? string.Empty).Trim();
                if (name.Length == 0 || string.Equals(name, code, StringComparison.OrdinalIgnoreCase)
                    || name.EndsWith($"({code})", StringComparison.OrdinalIgnoreCase))
                    ws.Cells[r, 2].Value = name.Length == 0 ? code : name;
                else
                    ws.Cells[r, 2].Value = $"{name} ({code})";
            }
            ws.DeleteColumn(22);
        }

        private static int LastAuditDataRow(ExcelWorksheet ws)
        {
            int last = 1;
            int end = ws.Dimension?.End.Row ?? 1;
            for (int r = 2; r <= end; r++)
            {
                if (AuditCodeAt(ws, r).Length > 0) last = r;
            }
            return last;
        }

        private static int GetOrAddAuditRow(ExcelWorksheet ws, string schoolCode)
        {
            string code = (schoolCode ?? string.Empty).Trim();
            int end = ws.Dimension?.End.Row ?? 1;
            for (int r = 2; r <= end; r++)
            {
                if (string.Equals(AuditCodeAt(ws, r), code, StringComparison.OrdinalIgnoreCase)) return r;
            }
            return LastAuditDataRow(ws) + 1;
        }

        private static void ClearAuditTotals(ExcelWorksheet ws)
        {
            int end = ws.Dimension?.End.Row ?? 1;
            for (int r = 2; r <= end; r++)
            {
                string label = (ws.Cells[r, 3].Text ?? string.Empty).Trim();
                if (label.StartsWith("Sub Total", StringComparison.OrdinalIgnoreCase)
                    || label.StartsWith("Grand Total", StringComparison.OrdinalIgnoreCase))
                {
                    ws.Cells[r, 1, r, AuditLastColumn].Clear();
                }
            }
        }

        // Writes Sub Total / Grand Total right below the last school row. Returns the grand total row.
        private static int WriteAuditTotals(ExcelWorksheet ws)
        {
            int lastData = LastAuditDataRow(ws);
            int subRow = lastData + 1;
            int grandRow = lastData + 2;
            ws.Cells[$"C{subRow}"].Value = "Sub Total: ";
            ws.Cells[$"C{subRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            foreach (char column in "DEFGHIJKLMNOPQR")
            {
                string cellAddress = $"{column}{subRow}";
                ws.Cells[cellAddress].Formula = $"SUM({column}2:{column}{lastData})";
                ws.Cells[cellAddress].Style.Numberformat.Format = "OPQR".Contains(column) ? "0" : "$#,##0.00";
            }
            ws.Cells[$"C{grandRow}"].Value = "Grand Total: ";
            ws.Cells[$"C{grandRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells[$"D{grandRow}"].Formula = $"=F{subRow}";

            // Sub Total and Grand Total rows: same back color as the data rows, bold text.
            foreach (int totalRow in new[] { subRow, grandRow })
            {
                using (ExcelRange range = ws.Cells[totalRow, 1, totalRow, 21])
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.MintCream);
                    range.Style.Font.Bold = true;
                    range.Style.Font.Color.SetColor(Color.Black);
                }
            }
            ws.Cells[$"D{grandRow}"].Style.Numberformat.Format = "$#,##0.00";

            // Yellow highlight on the Sub Total of: Total Amount (Dashboard) F, Last Payment using MM H, Service Fee Amount K.
            foreach (string col in new[] { "F", "H", "K" })
            {
                ws.Cells[$"{col}{subRow}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[$"{col}{subRow}"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            }

            // Acts Activated (Last 24 hrs) (column S) >= 1 is highlighted yellow.
            for (int r = 2; r <= lastData; r++)
            {
                if (decimal.TryParse(ws.Cells[r, 19].Text, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal activated)
                    && activated >= 1)
                {
                    ws.Cells[r, 19].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    ws.Cells[r, 19].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }
            }
            return grandRow;
        }

        public Task LogAuditIssueAsync(string message)
        {
            return File.AppendAllTextAsync(path, $"{message}{Environment.NewLine}");
        }
        #endregion

        #region Private Conversion Methods and Reusable methods

        private static void EnsureAuditHeaders(ExcelWorksheet workSheet)
        {
            if (workSheet.Dimension != null && workSheet.Dimension.Rows > 0)
            {
                MigrateLegacyCodeColumn(workSheet);
                return;
            }

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
                "Acts Activated (Last 24 hrs)",
                "New/ InProgress Account",
                "Funding Amount vs Transaction Amount (902 & 905)"
            };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = workSheet.Cells[1, i + 1];
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(Color.DarkOliveGreen);
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }
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
            await WaitForLoadingOverlayToDisappearAsync();
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
        // Makes sure the "Show alumni/disabled users" checkbox is ticked, verifying the state after each
        // click. Returns true if it had to tick it (so the caller knows the user list was reloaded).
        private async Task<bool> EnsureAlumniDisabledCheckedAsync()
        {
            await ChkAlumniDisableduser.First.WaitForAsync(
                new() { State = WaitForSelectorState.Visible, Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });

            var input = ChkAlumniDisableduser.First.Locator("input[type='checkbox']").First;
            var hasInput = await ChkAlumniDisableduser.First.Locator("input[type='checkbox']").CountAsync() > 0;
            Func<Task<bool>> isChecked = async () => hasInput && await input.IsCheckedAsync();

            if (!hasInput)
            {
                // No inspectable checkbox: keep the original single click behaviour.
                await ChkAlumniDisableduser.First.ClickAsync();
                await WaitForLoadingOverlayToDisappearAsync(timeout: MaximumUiWaitTimeoutMs);
                return true;
            }

            bool ticked = false;
            for (int attempt = 1; attempt <= 3 && !await isChecked(); attempt++)
            {
                ticked = true;
                try { await input.CheckAsync(new() { Force = true, Timeout = 5000 }); }
                catch (PlaywrightException) { await ChkAlumniDisableduser.First.ClickAsync(); }
                await WaitForLoadingOverlayToDisappearAsync(timeout: MaximumUiWaitTimeoutMs);
            }

            if (!await isChecked())
                throw new InvalidOperationException("The 'Show alumni/disabled users' checkbox could not be ticked on the ledger page.");
            return ticked;
        }

        // To find the total charges and payments posted for the given date range in School Ledger
        public async Task<LedgerPayments> GetChargesPaymentsAsync()
        {
            LedgerPayments ledgerPayments = new LedgerPayments();
            // "Show alumni/disabled users" must be ticked BEFORE choosing All / All Families,
            // because ticking it reloads the user/family list.
            await EnsureAlumniDisabledCheckedAsync();

            var userScope = _page.Locator("#userdiv:visible, #familydiv:visible");
            await userScope.WaitForAsync(
                new() { State = WaitForSelectorState.Visible, Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });
            bool isUser = await _page.Locator("#userdiv").IsVisibleAsync();
            if (isUser)
            {
                await SelectLedgerScopeAsync(DdlUser, SelectAllUsers, "(All)");
            }
            else
            {
                await SelectLedgerScopeAsync(DdlFamilyUser, SelectAll, "(All Families)");
            }

            await WaitForLoadingOverlayToDisappearAsync(timeout: MaximumUiWaitTimeoutMs);
            // A reload after the selection would untick the box and reset the list; confirm it is still ticked.
            if (await EnsureAlumniDisabledCheckedAsync())
            {
                if (await _page.Locator("#userdiv").IsVisibleAsync())
                    await SelectLedgerScopeAsync(DdlUser, SelectAllUsers, "(All)");
                else
                    await SelectLedgerScopeAsync(DdlFamilyUser, SelectAll, "(All Families)");
                await WaitForLoadingOverlayToDisappearAsync(timeout: MaximumUiWaitTimeoutMs);
            }
            await TxtReference.ClearAsync();

            // The ledger filters by calendar date, so include both dates touched by the previous 24 hours.
            DateTime today = DateTime.Today;
            DateTime yesterday = today.AddDays(-1);
            string startDate = yesterday.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            string endDate = today.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            await SetDateInputValueAsync(TxtStartDate, yesterday, "ledger start");
            await SetDateInputValueAsync(TxtEndDate, today, "ledger end");
            if (!IsDateValue(yesterday, await TxtStartDate.InputValueAsync()))
                await SetDateInputValueAsync(TxtStartDate, yesterday, "ledger start");
            if (!IsDateValue(today, await TxtEndDate.InputValueAsync()))
                await SetDateInputValueAsync(TxtEndDate, today, "ledger end");
            string selectedDate = await TxtStartDate.InputValueAsync();
            string selectedEndDate = await TxtEndDate.InputValueAsync();

            Console.WriteLine(
                $"Expected ledger date range: {startDate} - {endDate}, " +
                $"selected: {selectedDate} - {selectedEndDate}");

            await BtnLedgerFilter.WaitForAsync(
                new() { State = WaitForSelectorState.Visible, Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });
            if (!await BtnLedgerFilter.IsEnabledAsync())
            {
                throw new InvalidOperationException("The Ledger Filter button is disabled after selecting the user scope.");
            }

            await BtnLedgerFilter.ClickAsync();
            // Wait for the ledger table to be present
            var ledgerTableLocator = _page.Locator("#tblLedger");
            await WaitForLoadingOverlayToDisappearAsync();
            await ledgerTableLocator.WaitForAsync(new() { State = WaitForSelectorState.Visible });
            var ledgerTable = await ledgerTableLocator.IsVisibleAsync();
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

        private async Task SelectLedgerScopeAsync(ILocator dropdown, ILocator allOption, string expectedSelection)
        {
            var selectedScope = _page.Locator(
                $"//span[contains(@id, 'Users') and contains(normalize-space(.), '{expectedSelection}')]");

            // The select2 selection occasionally does not stick (or the page resets it); retry the selection.
            Exception lastError = null;
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    // Wait for the page/dropdown to finish loading BEFORE selecting.
                    await WaitForLoadingOverlayToDisappearAsync(timeout: MaximumUiWaitTimeoutMs);
                    await dropdown.WaitForAsync(
                        new() { State = WaitForSelectorState.Visible, Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });
                    try { await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 10000 }); } catch (TimeoutException) { }
                    await dropdown.ClickAsync();

                    // Wait until the select2 list is populated (not "Searching..."/"Loading...") before choosing.
                    var options = _page.Locator(".select2-results__option");
                    await options.First.WaitForAsync(
                        new() { State = WaitForSelectorState.Visible, Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });
                    await _page.WaitForFunctionAsync(
                        "() => { const o = document.querySelectorAll('.select2-results__option'); " +
                        "return o.length > 0 && ![...o].some(x => /searching|loading/i.test(x.textContent)); }",
                        null,
                        new() { Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });
                    await allOption.WaitForAsync(
                        new() { State = WaitForSelectorState.Visible, Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs });
                    await allOption.ClickAsync();

                    // Wait for the page to finish reloading data AFTER selecting.
                    await WaitForLoadingOverlayToDisappearAsync(timeout: MaximumUiWaitTimeoutMs);
                    try { await _page.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 10000 }); } catch (TimeoutException) { }

                    // Check the dropdown itself as well as the generic span.
                    await selectedScope.First.WaitForAsync(
                        new() { State = WaitForSelectorState.Visible, Timeout = 15000 });
                    return;
                }
                catch (Exception ex) when (ex is TimeoutException || ex is PlaywrightException)
                {
                    lastError = ex;
                    string current = string.Empty;
                    try { current = (await dropdown.InnerTextAsync(new() { Timeout = 3000 })).Trim(); } catch { }
                    if (current.Contains(expectedSelection, StringComparison.OrdinalIgnoreCase))
                    {
                        return; // selection is actually applied; the generic span just wasn't matched
                    }
                    Console.WriteLine($"Ledger scope '{expectedSelection}' not applied (attempt {attempt}, shows '{current}'); retrying.");
                    try { await _page.Keyboard.PressAsync("Escape"); } catch { }
                    await WaitForLoadingOverlayToDisappearAsync(timeout: MaximumUiWaitTimeoutMs);
                }
            }
            throw new InvalidOperationException(
                $"Could not select ledger scope '{expectedSelection}' after 3 attempts.", lastError);
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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            int lastRow = tab + 2;
            //workSheet[$"A{lastRow}"].Value = "";
            await Task.CompletedTask;
        }

        public async Task FinalViewAsync()
        {
            System.Diagnostics.Process.Start("notepad.exe", path);
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
            string ccAmount = "0";
            string eCheckAmount = "0";
            const string creditFooterSelector =
                "//*[@id='tblLedger_wrapper']/div[2]/div/div[3]/div/table/tfoot/tr[1]/td[1]";

            await DdlPaymentMethod.WaitForAsync();
            await DdlPaymentMethod.ClickAsync();
            await SelectCreditCard.WaitForAsync();
            await SelectCreditCard.ClickAsync();
            await BtnLedgerFilter.WaitForAsync();
            await BtnLedgerFilter.ClickAsync();
            await WaitForLoadingOverlayToDisappearAsync();
            ccAmount = await GetLedgerAmountAsync(creditFooterSelector, creditFooterSelector);

            await DdlPaymentMethod.ClickAsync();
            await SelecteCheck.WaitForAsync();
            await SelecteCheck.ClickAsync();
            await BtnLedgerFilter.ClickAsync();
            await WaitForLoadingOverlayToDisappearAsync();
            eCheckAmount = await GetLedgerAmountAsync(creditFooterSelector, creditFooterSelector);

            decimal totalMMPayment = ParseLedgerAmount(ccAmount) + ParseLedgerAmount(eCheckAmount);

            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine("Total Credit Card Payments(Last 24 hrs): " + ccAmount + "\r\n" + "Total eCheck Payments(Last 24 hrs): " + eCheckAmount + "\r\n" + "Total MM Payment(Last 24 hrs) : " + totalMMPayment + "\r\n");
            }
            ledgerPayments.CCPayment = ccAmount;
            ledgerPayments.eCheckPayment = eCheckAmount;
            ledgerPayments.TotalPayment = totalMMPayment.ToString();
            return ledgerPayments;
        }

        private static decimal ParseLedgerAmount(string amount)
        {
            if (decimal.TryParse(
                    amount,
                    NumberStyles.Currency,
                    CultureInfo.GetCultureInfo("en-US"),
                    out var parsedAmount))
            {
                return parsedAmount;
            }

            throw new FormatException($"Could not parse ledger payment amount '{amount}'.");
        }
        #endregion


        #region Verify School MM Dashboard and read and write
        public async Task<DashboardAmount> VerifySchoolMMDashboardAsync(string startDate)
        {
            DashboardAmount dashboardAmount = new DashboardAmount();
            //await _page.Locator("#tabMMDashboard").WaitForAsync();
            await TabMMDashboard.WaitForAsync(new() { State = WaitForSelectorState.Visible });
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
            await SetDateInputValueAsync(TxtStartDate, dateToSelect, "dashboard start");

            #endregion
            // Wait for next element
            await TxtEndDate.WaitForAsync(new()
            {
                Timeout = PlaywrightTimeoutPolicy.ActionTimeoutMs
            });

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
            await WaitForLoadingOverlayToDisappearAsync();
            string CCTotAmt = "0";
            string eCheckTotAmt = "0";
            decimal TotalAmount = 0;
            await _page.EvaluateAsync(
                "window.scrollBy(0, document.documentElement ? document.documentElement.scrollHeight : 0);");
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



