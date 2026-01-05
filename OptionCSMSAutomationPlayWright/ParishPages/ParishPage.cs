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
using OptionCSMSAutomationPlayWright.SISPages;
using OptionCSMSAutomationPlayWright.ParishPages;
using OpenQA.Selenium;
using CsvHelper;
using OpenQA.Selenium.BiDi.Modules.BrowsingContext;
using System.Globalization;
using static Microsoft.AspNetCore.Razor.Language.TagHelperMetadata;
using Dynamitey.DynamicObjects;

namespace OptionCSMSAutomationPlayWright.ParishPages
{
    public class ParishPage : BasePageObject
    {
        private IPage _page;
        private WaitHelper waitHelper;
        // Fixed folder path (read-only) for storing files or reports
        private readonly string directoryPath = "D:\\File";
        public string path = string.Empty;
        // Fixed file path for saving automation log details.
        private readonly string logFilePath = @"D:\File\ParishAutomationLog.txt";

        public ParishPage(IPage page, WaitHelper waitHelper) : base(page)
        {
            _page = page;
            _page.SetDefaultTimeout(100000);
            this.waitHelper = waitHelper;
        }

        #region Xpath Elements
        //Sign In Elements
        public ILocator TxtParishUserName => _page.Locator("//*[@id='txtEmail']");
        public ILocator TxtParishPassword => _page.Locator("//*[@id='txtPassword']");
        public ILocator ChkAuthorizedUser => _page.Locator("//*[@id='chkAuthorizedUser1']");
        public ILocator BtnParishLogin => _page.Locator("//*[@id='btnLogin']");
        //Parish Instruction Elements
        public ILocator ParishInstructions => _page.Locator("//span[normalize-space(text())='OptionC Parish Beta Testing Instructions']");
        //Saint of the Day Elements
        public ILocator SODSection => _page.Locator("//div[contains(@class,'db-panel-body') and contains(@class,'db-componenet-four')]").First;
        public ILocator SODReadMore => _page.Locator("(//a[normalize-space(text())='read more'])[1]");
        public ILocator SODGetSaintName => _page.Locator("//h4[@class='sectiontitle']");
        public ILocator SODGetSaintFeastDate => _page.Locator("//span[contains(normalize-space(text()),'Feast Day:')]");
        public ILocator ParishLogo => _page.Locator("//img[@alt='Parish Manager Logo']");
        //Daily Readings Elements
        public ILocator DailyReadingsSection => _page.Locator("//h3[contains(normalize-space(.), 'USCCB.org Daily Readings')]");
        public ILocator DRDate => _page.Locator("//h3[normalize-space()='USCCB.org Daily Readings']/following::span[contains(@style,'float:right')][1]");
        public ILocator DRReadMore => _page.Locator("//a[@href='/daily-readings' and normalize-space(text())='read more']");
        public ILocator DRTitleToday => _page.Locator("//div[contains(@class,'dr-title')]/h2");
        //Events Elements
        public ILocator EventsSection => _page.Locator("//h3[normalize-space()='Events']");
        public ILocator EventsList => _page.Locator("//div[contains(@class,'db-panel-body') and contains(@class,'db-componenet-two')]");
        //Acutis Elements
        //Sign In Elements
        public ILocator TxtAcutisUserName => _page.Locator("//*[@id='username']");
        public ILocator TxtAcutisPassword => _page.Locator("//*[@id='password']");     
        public ILocator BtnAcutisLogin => _page.Locator("//*[@id='Login']");
        public ILocator TicketsMenu => _page.Locator("(//*[@id='liTickets'])[1]");
        public ILocator RecentTicketsMenu => _page.Locator("(//*[@id='liRecentTickets'])[1]");
        private ILocator TicketRows => _page.Locator("//table[@id='ticketTable']//tbody//tr");
        private ILocator SubmittedDateCells => _page.Locator("//table[@id='ticketTable']//tbody//tr//td[3]");

        #endregion

        // 🔹 Common logging method with formatting
        private void LogToFile(string message, bool isSectionHeader = false)
        {
            // Ensure the log directory exists (create if missing)
            Directory.CreateDirectory(Path.GetDirectoryName(logFilePath)!);
            // Open the log file in append mode to add new entries without overwriting existing content
            using (StreamWriter writer = new StreamWriter(logFilePath, append: true))
            {
                if (isSectionHeader)
                {
                    // Add section header formatting with timestamp
                    writer.WriteLine();
                    // Writes a separator line (80 '═' characters) for better visual section division in the log file
                    writer.WriteLine(new string('═', 80));
                    writer.WriteLine($"📘 {message.ToUpper()}  |  {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    writer.WriteLine(new string('─', 80));
                }
                else
                {
                    // Write normal log entry with time
                    writer.WriteLine($"{DateTime.Now:HH:mm:ss}  {message}");
                }
            }

            // Also print to console
            Console.WriteLine(isSectionHeader
                ? $"\n📘 {message.ToUpper()}  |  {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n{new string('-', 80)}"
                : message);
        }

        // 🔹 Clears old log and creates a new log file with a header (optional)
        public void InitializeLog()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(logFilePath)!); // Ensure log folder exists
            File.WriteAllText(logFilePath, $"🕊️ Parish Dashboard Automation Log — {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n\n"); // Create/reset log file with header
        }


        //Methods for passing the user credentials
        // Load URL method    
        public async Task LoadURLAsync(string URL)
        {
            LogToFile("Initializing test session...", true);
            await _page.GotoAsync(URL, new() { Timeout = 100000 });
            LogToFile($"🌐 Loaded URL: {URL}");
        }
        //Enter the parish user credentials
        public async Task EnterParishCredentialsAsync(string username, string password)
        {
            LogToFile("Performing Parish Login...", true);
            await TxtParishUserName.FillAsync(username);
            await TxtParishPassword.FillAsync(password);
            await ChkAuthorizedUser.ClickAsync();
            await BtnParishLogin.ClickAsync();
            LogToFile($"🔑 Logged in with user: {username}");
        }
        //To maximize the window
        public async Task MaxWindow()
        {
            await _page.SetViewportSizeAsync(1500, 1080);
            await _page.EvaluateAsync("window.moveTo(0, 0); window.resizeTo(screen.width, screen.height);");
            LogToFile("🪟 Browser window maximized.");
        }
            //To verify the Parish Instruction in the dashboard
            public async Task VerifyAndGetParishInstructionsTextAsync()
        {
            LogToFile("Verifying Parish Instructions Section", true);
            string expectedText = "OptionC Parish Beta Testing Instructions";
            await ParishInstructions.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });

            if (!await ParishInstructions.IsVisibleAsync())
            {
                LogToFile("❌ Parish Instructions section not visible.");
                return;
            }

            string actualText = (await ParishInstructions.InnerTextAsync()).Trim();
            if (actualText == expectedText)
                LogToFile($"✅ Parish Instructions text matches expected value: '{actualText}'");
            else
                LogToFile($"⚠️ Mismatch: Expected '{expectedText}', but found '{actualText}'.");
        }

        //To check the Saint of the Day functionality
        public async Task VerifySaintOfTheDaySectionAsync()
        {
            LogToFile("Verifying Saint of the Day Section", true);
            await _page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

            if (!await SODSection.IsVisibleAsync())
            {
                LogToFile("❌ Saint of the Day section not visible.");
                return;
            }

            await SODReadMore.ClickAsync();
            await _page.WaitForLoadStateAsync();

            string saintName = (await SODGetSaintName.InnerTextAsync()).Trim();
            string feastDate = (await SODGetSaintFeastDate.InnerTextAsync()).Trim();
            LogToFile($"📜 Saint Name: {saintName}");
            LogToFile($"📅 Feast Date: {feastDate}");

            await ParishLogo.ClickAsync();
            await _page.WaitForLoadStateAsync();
            LogToFile("🏠 Returned to Parish Dashboard after viewing Saint details.");
        }
        //To check the daily readings functionality
        public async Task VerifyDailyReadingsOfTheDaySectionAsync()
        {
            LogToFile("Verifying Daily Readings Section", true);
            await _page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

            if (!await DailyReadingsSection.IsVisibleAsync())
            {
                LogToFile("❌ Daily Readings section not visible.");
                return;
            }

            string readingDate = (await DRDate.InnerTextAsync()).Trim();
            await DRReadMore.ClickAsync();
            await _page.WaitForLoadStateAsync();

            string readingTitle = (await DRTitleToday.InnerTextAsync()).Trim();
            LogToFile($"📖 Reading Title: {readingTitle}");
            LogToFile($"📅 Reading Date: {readingDate}");

            await ParishLogo.ClickAsync();
            await _page.WaitForLoadStateAsync();
            LogToFile("🏠 Returned to Parish Dashboard after verifying Daily Readings.");
        }

        public async Task VerifyEventsAtDashboardAsync()
        {
            LogToFile("Verifying Events Section", true);

            await _page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

            if (!await EventsSection.IsVisibleAsync())
            {
                LogToFile("❌ Events section not visible.");
                return;
            }

            // Wait for the <ul> list to appear
            var eventsList = _page.Locator("//h3[normalize-space()='Events']/following::ul[1]");
            await eventsList.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });

            // ✅ Get all <li> items inside that <ul>
            var eventItems = _page.Locator("//h3[normalize-space()='Events']/following::ul[1]/li");

            int count = await eventItems.CountAsync();

            if (count == 0)
                LogToFile("⚠️ No events are available in the list.");
            else
                LogToFile($"📅 Total Events Listed: {count}");
        }

        //Acutis Sign in
        // Load URL method    
        public async Task LoadAcutisURLAsync(string URL)
        {
            LogToFile("Initializing test session...", true);
            await _page.GotoAsync(URL, new() { Timeout = 10000 });
            LogToFile($"🌐 Loaded URL: {URL}");
        }
        //Enter the parish user credentials
        public async Task EnterAcutisParishCredentialsAsync(string Username, string Password)
        {
           LogToFile("Performing Acutis Parish Login...", true);
            await TxtAcutisUserName.FillAsync(Username);
            await TxtAcutisPassword.FillAsync(Password);
            await BtnAcutisLogin.ClickAsync();
            LogToFile($"🔑 Logged in with user: {Username}");
        }
       
       
        public async Task NavigateToRecentTicketsAsync()
        {
            await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
            await TicketsMenu.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
            await RecentTicketsMenu.ClickAsync();
            await _page.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
        }
        // ✅ Verify total ticket count
        public async Task VerifyParishTicketsAsync()
        {
            // Wait until table is visible
            await _page.WaitForSelectorAsync("//table[@id='tblBilling']/tbody/tr");

            // Get ticket row count
            var ticketRows = await _page.Locator("//table[@id='tblBilling']/tbody/tr").CountAsync();

            // Log the count
            LogToFile($"📋 Total Tickets Count: {ticketRows}");
        }

        // ✅ Get latest submitted ticket details(US date format)
        public async Task<(string TicketId, string ParishId, string ParishName, string SubmittedBy, string SubmittedDate)> GetLatestTicketDetailsAsync()
        {
            await _page.WaitForSelectorAsync("//table[@id='tblBilling']/tbody/tr");

            // ✅ Extract all date strings from the 11th column
            var dateStrings = await _page.Locator("//table[@id='tblBilling']/tbody/tr/td[11]").AllInnerTextsAsync();

            var usCulture = new CultureInfo("en-US");

            // ✅ Parse all valid US-format dates
            var parsedDates = dateStrings
                .Select(d =>
                {
                    if (DateTime.TryParseExact(d.Trim(), "MM/dd/yyyy hh:mm:ss tt", usCulture,
                        DateTimeStyles.None, out var dt))
                        return dt;
                    return DateTime.MinValue;
                })
                .Where(d => d != DateTime.MinValue)
                .ToList();

            if (parsedDates.Count == 0)
            {
                LogToFile("⚠️ No valid submitted dates found in table.");
                return ("", "", "", "", "");
            }

            // ✅ Find the latest date and convert to US format for matching
            var latestDate = parsedDates.Max();
            var latestDateString = latestDate.ToString("MM/dd/yyyy hh:mm:ss tt", usCulture);

            // ✅ Use partial date match in 11th column
            var latestRow = _page.Locator(
                $"//table[@id='tblBilling']/tbody/tr[td[11][contains(normalize-space(.), '{latestDate:MM/dd/yyyy}')]]"
            ).First;

            await latestRow.WaitForAsync(new LocatorWaitForOptions { Timeout = 15000 });

            // ✅ Extract columns using nth-child (safe CSS format)
            string ticketId = (await latestRow.Locator("td:nth-child(1)").InnerTextAsync()).Trim();
            string parishId = (await latestRow.Locator("td:nth-child(2)").InnerTextAsync()).Trim();
            string parishName = (await latestRow.Locator("td:nth-child(3)").InnerTextAsync()).Trim();
            string submittedBy = (await latestRow.Locator("td:nth-child(8)").InnerTextAsync()).Trim();

            // ✅ Log in the same U.S. date format
            LogToFile("🕒 Latest Submitted Ticket Details:");
            LogToFile($"📄 Ticket ID     : {ticketId}");
            LogToFile($"🏫 Parish ID     : {parishId}");
            LogToFile($"📖 Parish Name   : {parishName}");
            LogToFile($"📅 Submitted Date: {latestDate.ToString("MM/dd/yyyy hh:mm:ss tt", usCulture)}");
            LogToFile($"👤 Submitted By  : {submittedBy}");

            return (ticketId, parishId, parishName, submittedBy, latestDate.ToString("MM/dd/yyyy hh:mm:ss tt", usCulture));
        }


        public async Task VerifyTicketsListDisplayedAsync()
        {
            await _page.WaitForSelectorAsync("//table[@id='tblBilling']/tbody/tr");
            var isVisible = await _page.Locator("//table[@id='tblBilling']").IsVisibleAsync();

            if (isVisible)
                LogToFile("✅ Ticket list is displayed successfully.");
            else
                LogToFile("❌ Ticket list is not visible on the Recent Tickets page.");
        }

    }

}
