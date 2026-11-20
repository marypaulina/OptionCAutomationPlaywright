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
            int total = auditTable.Rows.Count;
            var failedRows = new List<int>();
            bool aborted = false;

            // Pass 1: every school
            for (int i = 0; i < total; i++)
            {
                if (!await AuditSchoolAsync(auditTable, i, total, null))
                {
                    failedRows.Add(i);
                }

                if (total != i + 1 && !await TryBackToAcutisAsync(auditTable.Rows[i]["SchoolCode"], failures))
                {
                    aborted = true;
                    break;
                }
            }

            // Pass 2: only the schools that failed in pass 1
            if (failedRows.Count > 0 && !aborted)
            {
                await mmPage.LogAuditIssueAsync(
                    $"RETRY PASS - re-running {failedRows.Count} failed school(s): " +
                    string.Join(", ", failedRows.Select(r => auditTable.Rows[r]["SchoolCode"])));

                if (await TryBackToAcutisAsync("retry start", failures))
                {
                    for (int k = 0; k < failedRows.Count; k++)
                    {
                        int i = failedRows[k];
                        var retryFailures = new List<Exception>();
                        if (!await AuditSchoolAsync(auditTable, i, total, retryFailures))
                        {
                            failures.AddRange(retryFailures); // still failing after retry
                        }

                        if (k != failedRows.Count - 1 &&
                            !await TryBackToAcutisAsync(auditTable.Rows[i]["SchoolCode"], failures))
                        {
                            break;
                        }
                    }
                }
            }
            else if (failedRows.Count > 0)
            {
                failures.Add(new InvalidOperationException(
                    "Retry pass skipped because the browser could not return to Acutis."));
            }

            if (failures.Count > 0)
            {
                throw new AggregateException(
                    $"The school audit completed with {failures.Count} error(s) after the retry pass. Text log: {mmPage.path}",
                    failures);
            }
        }

        // Audits one school and writes its Excel row / text log. Returns false when it failed.
        private async Task<bool> AuditSchoolAsync(Table auditTable, int i, int total, List<Exception>? failureSink)
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
                if (schoolIdentifier != "4291" && schoolIdentifier != "5940" && schoolIdentifier != "8518")
                {
                    dashboardAmount = await mmPage.VerifySchoolMMDashboardAsync(startDate);// Open dashboard, filter current school year, and read data
                }
                await mmPage.VerifyFeeManagementAsync(); // Open the Fee Management page
                objreport = await mmPage.LoopFilterAsync(schoolIdentifier);// Loop through all MM reports and get values
                await mmPage.VerifyAccountStatusAsync(objreport);// Track account status
                await mmPage.WriteAuditReportInExcel(objreport, ledgerPayments, dashboardAmount, i, startDate, total);// Write values to Excel
                return true;
            }
            catch (Exception exception)
            {
                var schoolFailure = new InvalidOperationException(
                    $"Audit failed for school '{schoolIdentifier}' (start date {startDate}).",
                    exception);
                Console.Error.WriteLine(schoolFailure);
                var sink = failureSink ?? new List<Exception>();
                sink.Add(schoolFailure);

                try
                {
                    await mmPage.RecordSchoolAuditFailureAsync(i, total, schoolIdentifier, startDate, exception);
                }
                catch (Exception recordException)
                {
                    sink.Add(new InvalidOperationException(
                        $"Could not record the failure for school '{schoolIdentifier}' in the audit outputs.",
                        recordException));
                }
                return false;
            }
        }

        private async Task<bool> TryBackToAcutisAsync(string context, List<Exception> failures)
        {
            try
            {
                await mmPage.BackToAcutisAsync();
                return true;
            }
            catch (Exception recoveryException)
            {
                string recoveryMessage = $"Could not return to Acutis after '{context}'; remaining schools were not run.";
                failures.Add(new InvalidOperationException(recoveryMessage, recoveryException));
                Console.Error.WriteLine(recoveryMessage);
                try
                {
                    await mmPage.LogAuditIssueAsync($"{recoveryMessage} {recoveryException.Message}");
                }
                catch (Exception logException)
                {
                    failures.Add(new InvalidOperationException("Could not log the audit issue after recovery failed.", logException));
                }
                return false;
            }
        }
    }
}
