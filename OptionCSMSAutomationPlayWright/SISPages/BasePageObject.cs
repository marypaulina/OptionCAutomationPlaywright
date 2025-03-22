using Atata;
using Microsoft.Playwright;

namespace OptionCSMSAutomationPlayWright.SISPages
{

    public abstract class BasePageObject
    {
        protected readonly IPage _page;
        private readonly WaitHelper waitHelper;

        public BasePageObject(IPage page)
        {
            _page = page;
            waitHelper = new WaitHelper(_page);
        }
        protected ILocator SearchInput => _page.Locator("//input[@placeholder='Search']");

        public async Task ClickAsync(string selector)
        {
            await _page.Locator(selector).ClickAsync();
        }

        public async Task FillAsync(string selector, string text)
        {
            await _page.Locator(selector).FillAsync(text);
        }

        public async Task<string> GetTextAsync(string selector)
        {
            return await _page.Locator(selector).InnerTextAsync();
        }

        public async Task WaitForElementAsync(string selector, int timeout = 50000)
        {
            await _page.Locator(selector).WaitForAsync(new LocatorWaitForOptions { Timeout = timeout });
        }

        public async Task<bool> IsElementVisibleAsync(string selector)
        {
            return await _page.Locator(selector).IsVisibleAsync();
        }

        public class WaitHelper
        {
            private readonly IPage _page;

            public WaitHelper(IPage page)
            {
                _page = page;
            }

            public async Task WaitForElementToBeVisible(ILocator locator)
            {
                await locator.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });
            }

            public async Task WaitForElementToBeEnabled(ILocator locator)
            {
                await locator.WaitForAsync(new LocatorWaitForOptions { State = WaitForSelectorState.Visible });
            }
        }
        // Class-level flag to track if paging has been set
        //private static bool isPagingSet = false;

        // Method to search by name
        public async Task SearchByNameAsync(string searchText)
        {
            try
            {
                await waitHelper.WaitForElementToBeVisible(SearchInput);
                await SearchInput.FillAsync(searchText);
                await Task.Delay(1000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        #region Date & Time

        // Common method for converting the given date into MM/DD/YYYY format
        public string DateFunction(string commonDate)
        {
            DateTime date = string.IsNullOrEmpty(commonDate) ? DateTime.Today : Convert.ToDateTime(commonDate);
            return date.ToString("MM/dd/yyyy");
        }

        // Common method for converting the given time into HH:MM TT format
        public string TimeFunction(string commonTime)
        {
            DateTime time = string.IsNullOrEmpty(commonTime) ? DateTime.Now : Convert.ToDateTime(commonTime);
            return time.ToString("hh:mm tt");
        }

        #endregion
    }
}