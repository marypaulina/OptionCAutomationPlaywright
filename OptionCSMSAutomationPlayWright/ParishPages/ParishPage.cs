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

        public ILocator LnkParishes => _page.Locator("(//a[@id='liSchool'])[1]");
        public ILocator TxtSearch => _page.Locator("//input[@type='search' and @aria-controls='example']");
        public ILocator BtnNextGen => _page.Locator("//a[normalize-space()='N']");

        // ================= Events Section Locators =================

        public ILocator EventsHeader =>
            _page.Locator("//h3[normalize-space()='Events']");

        public ILocator NoEventsMessage =>
            _page.Locator("//div[contains(@class,'norecord_style') and normalize-space()='No events available.']");

        public ILocator EventItems =>
            _page.Locator("//h3[normalize-space()='Events']/ancestor::div[contains(@class,'db-panel')]//div[contains(@class,'always-visible')]//div[not(contains(@class,'norecord_style')) and not(contains(@class,'ps-scrollbar'))]");

        // ================= Announcements Section Locators =================

        public ILocator AnnouncementsHeader =>
            _page.Locator("//h3[normalize-space()='Announcements']");

        public ILocator NoAnnouncementsMessage =>
            _page.Locator("//div[contains(@class,'norecord_style') and normalize-space()='No announcements available.']");

        public ILocator AnnouncementItems =>
            _page.Locator("//h3[normalize-space()='Announcements']/ancestor::div[contains(@class,'db-panel')]//div[contains(@class,'always-visible')]//div[not(contains(@class,'norecord_style')) and not(contains(@class,'ps-scrollbar'))]");
        // ================= New Messages Section Locators =================

        public ILocator NewMessagesHeader =>
            _page.Locator("//h3[normalize-space()='New Messages']");

        public ILocator NoNewMessagesMessage =>
            _page.Locator("//div[contains(@class,'norecord_style') and normalize-space()='No new messages available.']");

        public ILocator NewMessageItems =>
            _page.Locator("//h3[normalize-space()='New Messages']/ancestor::div[contains(@class,'db-panel')]//div[contains(@class,'always-visible')]//div[not(contains(@class,'norecord_style')) and not(contains(@class,'ps-scrollbar'))]");


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

        public async Task OpenTestParish()
        {
            try
            {
                LogToFile("OpenTestParish started");

                var oldPage = _page; // existing tab

                await LnkParishes.ClickAsync();
                LogToFile("Clicked Parishes link");

                await TxtSearch.FillAsync("15057");
                LogToFile("Entered parish search value");

                // Start waiting for new tab BEFORE clicking
                var waitForNewPage = _page.Context.WaitForPageAsync();

                await BtnNextGen.ClickAsync();
                LogToFile("Clicked NextGen button");

                // Capture new tab
                var newPage = await waitForNewPage;
                await newPage.WaitForLoadStateAsync(LoadState.DOMContentLoaded);
                LogToFile("New tab loaded");

                // Close old tab
                await oldPage.CloseAsync();
                LogToFile("Previous tab closed");

                // Switch context to new tab
                _page = newPage;
                LogToFile("Switched control to new tab");
            }
            catch (Exception ex)
            {
                LogToFile($"ERROR in OpenTestParish: {ex.Message}");
                throw;
            }
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

        // To check the Events section in Dashboard
        public async Task<int> VerifyEventsAtDashboardAsync()
        {
            LogToFile("Verifying Events Section at Dashboard", true);

            await _page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

            if (!await EventsHeader.IsVisibleAsync())
            {
                LogToFile("❌ Events section is not visible.");
                return 0;
            }

            if (await NoEventsMessage.IsVisibleAsync())
            {
                LogToFile("⚠️ No events are available in the list.");
                return 0;
            }

            int eventCount = await EventItems.CountAsync();

            LogToFile($"✅ Total Events Available: {eventCount}");

            for (int i = 0; i < eventCount; i++)
            {
                string eventText = (await EventItems.Nth(i).InnerTextAsync()).Trim();
                LogToFile($"   ➤ Event {i + 1}: {eventText}");
            }

            LogToFile("✔️ Events section verification completed successfully.");

            return eventCount;
        }



        // To check the Task List functionality
        public async Task VerifyTaskListSectionAsync()
        {
            LogToFile("Verifying Task List Section", true);

            await _page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

            var taskListSection = _page.Locator("//h3[normalize-space()='Task List']");

            if (!await taskListSection.IsVisibleAsync())
            {
                LogToFile("❌ Task List section is not visible.");
                return;
            }

            var taskLinks = _page.Locator(
                "//h3[normalize-space()='Task List']/ancestor::div[contains(@class,'db-panel')]//ul/li/a"
            );

            int taskCount = await taskLinks.CountAsync();

            if (taskCount == 0)
            {
                LogToFile("⚠️ No tasks are available in the Task List.");
                return;
            }

            LogToFile($"✅ Total Tasks Available: {taskCount}");

            for (int i = 0; i < taskCount; i++)
            {
                string taskName = (await taskLinks.Nth(i).InnerTextAsync()).Trim();
                LogToFile($"   ➤ Task {i + 1}: {taskName}");
            }

            LogToFile("✔️ Task List verification completed successfully.");
        }

        // To check the Announcements section in Dashboard
        public async Task<int> VerifyAnnouncementsAtDashboardAsync()
        {
            LogToFile("Verifying Announcements Section at Dashboard", true);

            await _page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

            if (!await AnnouncementsHeader.IsVisibleAsync())
            {
                LogToFile("❌ Announcements section is not visible.");
                return 0;
            }

            if (await NoAnnouncementsMessage.IsVisibleAsync())
            {
                LogToFile("⚠️ No announcements are available in the list.");
                return 0;
            }

            int announcementCount = await AnnouncementItems.CountAsync();

            LogToFile($"✅ Total Announcements Available: {announcementCount}");

            for (int i = 0; i < announcementCount; i++)
            {
                string announcementText = (await AnnouncementItems.Nth(i).InnerTextAsync()).Trim();
                LogToFile($"   ➤ Announcement {i + 1}: {announcementText}");
            }

            LogToFile("✔️ Announcements section verification completed successfully.");

            return announcementCount;
        }

        // To check the New Messages section in Dashboard
        public async Task<int> VerifyNewMessagesAtDashboardAsync()
        {
            LogToFile("Verifying New Messages Section at Dashboard", true);

            await _page.EvaluateAsync("window.scrollTo(0, document.body.scrollHeight)");

            if (!await NewMessagesHeader.IsVisibleAsync())
            {
                LogToFile("❌ New Messages section is not visible.");
                return 0;
            }

            if (await NoNewMessagesMessage.IsVisibleAsync())
            {
                LogToFile("⚠️ No new messages are available in the list.");
                return 0;
            }

            int messageCount = await NewMessageItems.CountAsync();

            LogToFile($"✅ Total New Messages Available: {messageCount}");

            for (int i = 0; i < messageCount; i++)
            {
                string messageText = (await NewMessageItems.Nth(i).InnerTextAsync()).Trim();
                LogToFile($"   ➤ Message {i + 1}: {messageText}");
            }

            LogToFile("✔️ New Messages section verification completed successfully.");

            return messageCount;
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
