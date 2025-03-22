using Microsoft.Playwright;
using OptionCSMSAutomationPlayWright.Model;
using OptionCSMSAutomationPlayWright.SISPages;
using TechTalk.SpecFlow;
using static OptionCSMSAutomationPlayWright.SISPages.BasePageObject;
using Table = TechTalk.SpecFlow.Table;

namespace OptionCSMSAutomationPlayWright.StepDefinitions
{
    [Binding]
    public sealed class MMAuditStepDefinitions
    {
        private readonly IPage page; // Using Playwright's IPage instead of IWebDriver
        private readonly MMAuditPages mmPage;
        //private readonly WaitHelper objWaitHelper;
        public MMAuditStepDefinitions(IPage page, WaitHelper objWaitHelper) // Constructor
        {
            this.page = page;
            //this.objWaitHelper = objWaitHelper;
            this.mmPage = new MMAuditPages(page, objWaitHelper);

        }

        [Given(@"Acutis User has successfully launched")]
        public async Task GivenAcutisUserHasSuccessfullyLaunched(Table userCredentialsTable)
        {

            if (userCredentialsTable != null && userCredentialsTable.Rows.Count > 0)
            {
                await mmPage.LoadURLAsync(userCredentialsTable.Rows[0]["URL"]); //This is to load the URL
                await mmPage.EnterAcutisCredentialsAsync(userCredentialsTable.Rows[0]["Username"], userCredentialsTable.Rows[0]["Password"]); //This is to enter the acutis user credentials
            }
        }

        [Given(@"Open all the MM schools and audit the fee details everyday")]
        public async Task GivenOpenAllTheMMSchoolsAndAuditTheFeeDetailsEveryday(Table auditTable)
        {
            for (int i = 0; i < auditTable.Rows.Count; i++)
            {
                Report objreport = new Report();
                LedgerPayments ledgerPayments = new LedgerPayments();
                DashboardAmount dashboardAmount = new DashboardAmount();

                string schoolCode = auditTable.Rows[i]["SchoolCode"];
                string startDate = auditTable.Rows[i]["StartDate"];
                await mmPage.OpenSchoolAsync(auditTable.Rows[i]["SchoolCode"]); // Search and open the school
                await mmPage.NavigateToFeeManagementAsync(); // Navigate to Fee Management
                Console.WriteLine("Landed in Fee Management page");
                ledgerPayments = await mmPage.GetChargesPaymentsAsync();// Verify charges, payments posted in last 24 hrs from the ledger
                if (schoolCode != "3932" && schoolCode != "4291")
                {
                    dashboardAmount = await mmPage.VerifySchoolMMDashboardAsync(startDate);// Open dashboard, filter current school year, and read data
                }
                await mmPage.VerifyFeeManagementAsync(); // Open the Fee Management page                
                objreport = await mmPage.LoopFilterAsync();// Loop through all MM reports and get values               
                await mmPage.VerifyAccountStatusAsync(objreport, page);// Track account status                
                await mmPage.WriteAuditReportInExcel(objreport, ledgerPayments, dashboardAmount, i, startDate, auditTable.Rows.Count);// Write values to Excel               
                if (auditTable.Rows.Count != i + 1)// Go back to Acutis after auditing one school
                {
                    await mmPage.BackToAcutisAsync();
                }
            }
        }
    }
}