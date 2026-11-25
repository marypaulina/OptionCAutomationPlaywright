using Microsoft.Playwright;
using OptionCSMSAutomationPlayWright.SISPages;
using TechTalk.SpecFlow;

namespace OptionCSMSAutomationPlayWright.StepDefinitions
{
    [Binding]
    public sealed class AgingReportStepDefinitions
    {
        private readonly AgingReportPage agingPage;

        public AgingReportStepDefinitions(IPage page)
        {
            agingPage = new AgingReportPage(page);
        }

        [Given(@"Open the dynamic report ""(.*)""")]
        public async Task GivenOpenTheDynamicReport(string url) => await agingPage.OpenReportAsync(url);

        [When(@"Display all the records using the paging option")]
        public async Task WhenDisplayAllTheRecords() => await agingPage.ShowAllRecordsAsync();

        [When(@"Sort the report by Aging Days from the largest to the smallest")]
        public async Task WhenSortByAgingDaysDescending() => await agingPage.SortByAgingDaysDescendingAsync();

        [When(@"Highlight the Aging Days greater than (.*)")]
        public async Task WhenHighlightAgingDaysGreaterThan(int limit) => await agingPage.HighlightAgingDaysAboveAsync(limit);

        // ---- eCheck Service Fee Report ----

        [When(@"Go to the Reports menu, eCheck Report and open the Service Fee Report tab")]
        public async Task WhenOpenServiceFeeReport()
        {
            await agingPage.OpenReportsMenuItemAsync("liEcheckReports", "eCheck Report");
            await agingPage.OpenServiceFeeReportTabAsync();
        }

        [When(@"Select the Processing Start Date and Processing End Date as yesterday \(Friday to Sunday of the previous week on a Monday\) and click Run Report")]
        public async Task WhenFilterServiceFeeReport() => await agingPage.FilterServiceFeeReportAsync();

        [When(@"Resize the page to take the screenshot")]
        public async Task WhenResizeThePage() => await agingPage.ResizePageToFitReportAsync();

        [When(@"Highlight the Total data in yellow with a thick border")]
        public async Task WhenHighlightTotal() => await agingPage.HighlightServiceFeeTotalAsync();

        [When(@"Display the text ""Service Fee to OptionC on the selected date range"" in big bold red font")]
        public async Task WhenDisplayServiceFeeText() => await agingPage.DisplayServiceFeeTextAsync();

        [Then(@"Attach the Service Fee Report screenshot in the Aging Report sheet")]
        public async Task ThenAttachServiceFeeScreenshot() => await agingPage.AttachServiceFeeScreenshotAsync();

        // ---- Auto Withdrawals ----

        [When(@"Go to the Reports menu and open Auto Withdrawals")]
        public async Task WhenOpenAutoWithdrawals() => await agingPage.OpenReportsMenuItemAsync("liAutoWithdrawal", "Auto Withdrawals");

        [When(@"Select the current date as Start Date and End Date and click Filter")]
        public async Task WhenFilterAutoWithdrawals() => await agingPage.FilterAutoWithdrawalsForTodayAsync();

        [When(@"Display the count of auto withdrawal transactions scheduled today, or that none are scheduled")]
        public async Task WhenDisplayAutoWithdrawalText() => await agingPage.DisplayAutoWithdrawalTextAsync();

        [Then(@"Attach the Auto Withdrawals screenshot in the Aging Report sheet")]
        public async Task ThenAttachAutoWithdrawalScreenshot() => await agingPage.AttachAutoWithdrawalScreenshotAsync();

        [Then(@"Take a full screenshot of the report and insert it in the next sheet of the audit workbook for Aging Days greater than (.*)")]
        public async Task ThenInsertScreenshotInNextSheet(int limit) => await agingPage.InsertScreenshotInNextSheetAsync(limit);
    }
}
