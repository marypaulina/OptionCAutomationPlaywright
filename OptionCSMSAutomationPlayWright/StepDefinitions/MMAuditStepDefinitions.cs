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
            var failures = new List<Exception>();
            for (int i = 0; i < auditTable.Rows.Count; i++)
            {
                Report objreport = new Report();
                LedgerPayments ledgerPayments = new LedgerPayments();
                DashboardAmount dashboardAmount = new DashboardAmount();

                string schoolIdentifier = auditTable.Rows[i]["SchoolCode"];
                string startDate = auditTable.Rows[i]["StartDate"];
                try
                {
                    await mmPage.OpenSchoolAsync(schoolIdentifier); // Search and open the school
                    await mmPage.HandleMMMaintenancePopupIfPresentAsync();
                    await mmPage.HandleMMReactivatedPopupIfPresentAsync();
                    await mmPage.NavigateToFeeManagementAsync(); // Navigate to Fee Management
                    Console.WriteLine("Landed in Fee Management page");
                    ledgerPayments = await mmPage.GetChargesPaymentsAsync();// Verify charges, payments posted in last 24 hrs from the ledger
                    //if (schoolCode != "3932" && schoolCode != "4291")
                    if (schoolIdentifier != "4291" && schoolIdentifier != "5940" && schoolIdentifier != "8518")
                    {
                        dashboardAmount = await mmPage.VerifySchoolMMDashboardAsync(startDate);// Open dashboard, filter current school year, and read data
                    }
                    await mmPage.VerifyFeeManagementAsync(); // Open the Fee Management page
                    objreport = await mmPage.LoopFilterAsync(schoolIdentifier);// Loop through all MM reports and get values
                    await mmPage.VerifyAccountStatusAsync(objreport);// Track account status
                    await mmPage.WriteAuditReportInExcel(objreport, ledgerPayments, dashboardAmount, i, startDate, auditTable.Rows.Count);// Write values to Excel
                }
                catch (Exception exception)
                {
                    var schoolFailure = new InvalidOperationException(
                        $"Audit failed for school '{schoolIdentifier}' (start date {startDate}).",
                        exception);
                    failures.Add(schoolFailure);
                    Console.Error.WriteLine(schoolFailure);

                    try
                    {
                        await mmPage.RecordSchoolAuditFailureAsync(
                            i,
                            auditTable.Rows.Count,
                            schoolIdentifier,
                            startDate,
                            exception);
                    }
                    catch (Exception recordException)
                    {
                        failures.Add(new InvalidOperationException(
                            $"Could not record the failure for school '{schoolIdentifier}' in the audit outputs.",
                            recordException));
                    }

                    try
                    {
                        await mmPage.CreateAuditOutputCheckpointAsync();
                    }
                    catch (Exception checkpointException)
                    {
                        failures.Add(new InvalidOperationException(
                            $"Could not checkpoint the audit outputs after school '{schoolIdentifier}' failed.",
                            checkpointException));
                    }
                }

                if (auditTable.Rows.Count != i + 1)
                {
                    try
                    {
                        await mmPage.BackToAcutisAsync();
                    }
                    catch (Exception recoveryException)
                    {
                        string recoveryMessage =
                            $"Could not return to Acutis after school '{schoolIdentifier}'; remaining schools were not run.";
                        failures.Add(new InvalidOperationException(recoveryMessage, recoveryException));
                        Console.Error.WriteLine(recoveryMessage);

                        try
                        {
                            await mmPage.LogAuditIssueAsync($"{recoveryMessage} {recoveryException.Message}");
                            await mmPage.CreateAuditOutputCheckpointAsync();
                        }
                        catch (Exception checkpointException)
                        {
                            failures.Add(new InvalidOperationException(
                                "Could not log or checkpoint the audit after recovery failed.",
                                checkpointException));
                        }
                        break;
                    }
                }
            }

            if (failures.Count > 0)
            {
                throw new AggregateException(
                    $"The school audit completed with {failures.Count} error(s). Text log: {mmPage.path}",
                    failures);
            }
        }
    }
}