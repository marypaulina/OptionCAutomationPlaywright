using Microsoft.Playwright;
using OptionCSMSAutomationPlayWright.Model;
using OptionCSMSAutomationPlayWright.SISPages;
using TechTalk.SpecFlow;
using static OptionCSMSAutomationPlayWright.SISPages.BasePageObject;
using OptionCSMSAutomationPlayWright.ParishPages;
using Table = TechTalk.SpecFlow.Table;
using NUnit.Framework;

namespace OptionCSMSAutomationPlayWright.StepDefinitions
{
    [Binding]
    public sealed class ParishBasicFlowStepDefinitions
    {
         private readonly IPage page; // Using Playwright's IPage instead of IWebDriver
        private readonly ParishPage parishPage;
        //private readonly WaitHelper objWaitHelper;
        public ParishBasicFlowStepDefinitions(IPage page, WaitHelper objWaitHelper) // Constructor
        {
            this.page = page;
            //this.objWaitHelper = objWaitHelper;
            this.parishPage = new ParishPage(page, objWaitHelper);

        }
        //To sign in into parish
        [Given(@"Parish Admin user has succesfully launched")]
        public async Task GivenParishAdminUserHasSuccesfullyLaunched(Table userCredentialsTable)
        {
            await parishPage.LoadURLAsync(userCredentialsTable.Rows[0]["URL"]); //This is to load the URL
            await parishPage.EnterParishCredentialsAsync(userCredentialsTable.Rows[0]["Username"], userCredentialsTable.Rows[0]["Password"]); //This is to enter the acutis user credentials
        }

        [Given(@"I Verify whether all the parish dasboard elements are displayed for the parish admin user")]
        public async Task GivenIVerifyWhetherAllTheParishDasboardElementsAreDisplayedForTheParishAdminUser()
        {
            await parishPage.OpenTestParish();
            await parishPage.MaxWindow();
            await parishPage.VerifyAndGetParishInstructionsTextAsync();
            await parishPage.VerifySaintOfTheDaySectionAsync();
            await parishPage.VerifyDailyReadingsOfTheDaySectionAsync();
            await parishPage.VerifyEventsAtDashboardAsync();
            await parishPage.VerifyTaskListSectionAsync();
            await parishPage.VerifyAnnouncementsAtDashboardAsync();
            await parishPage.VerifyNewMessagesAtDashboardAsync();

        }

        //Acutis Step Definitions
        [Given(@"I sign in using Parish Acutis user credentials")]
        public async Task GivenISignInUsingParishAcutisUserCredentials(Table acutisCredentialsTable)
        {
            // Extract credentials from Gherkin table
            await parishPage.LoadAcutisURLAsync(acutisCredentialsTable.Rows[0]["URL"]); //This is to load the URL
            await parishPage.EnterAcutisParishCredentialsAsync(acutisCredentialsTable.Rows[0]["Username"], acutisCredentialsTable.Rows[0]["Password"]);
        }

        [When(@"I navigate to the Recent Tickets page")]
        public async Task WhenINavigateToTheRecentTicketsPage()
        {
            await parishPage.NavigateToRecentTicketsAsync();
        }

        [Then(@"I should see the list of tickets displayed")]
        public async Task ThenIShouldSeeTheListOfTicketsDisplayed()
        {
            await parishPage.VerifyTicketsListDisplayedAsync();
        }

        [Then(@"I should display the total count of tickets")]
        public async Task ThenIShouldDisplayTheTotalCountOfTickets()
        {
            await parishPage.VerifyParishTicketsAsync();
        }

        [Then(@"I should get the latest submitted ticket details including Ticket ID, School ID, School Name, Submitted By, and Submitted Date")]
        public async Task ThenIShouldGetTheLatestSubmittedTicketDetailsIncludingTicketIDSchoolIDSchoolNameSubmittedByAndSubmittedDate()
        {
            var details = await parishPage.GetLatestTicketDetailsAsync();

            Console.WriteLine($"Ticket ID: {details.TicketId}");
            Console.WriteLine($"Parish ID: {details.ParishId}");
            Console.WriteLine($"Parish Name: {details.ParishName}");
            Console.WriteLine($"Submitted By: {details.SubmittedBy}");
            Console.WriteLine($"Submitted Date (converted): {details.SubmittedDate}");
        }

    }
}
