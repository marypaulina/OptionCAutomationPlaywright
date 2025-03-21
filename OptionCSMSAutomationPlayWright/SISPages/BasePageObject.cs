using Microsoft.Playwright;
using System.Threading.Tasks;

namespace OptionCSMSAutomationPlayWright.SISPages
{

    public abstract class BasePageObject
    {
        protected readonly IPage Page;
        //private readonly WaitH waitHelper;

        public BasePageObject(IPage page)
        {
            Page = page;
        }

        public async Task ClickAsync(string selector)
        {
            await Page.Locator(selector).ClickAsync();
        }

        public async Task FillAsync(string selector, string text)
        {
            await Page.Locator(selector).FillAsync(text);
        }

        public async Task<string> GetTextAsync(string selector)
        {
            return await Page.Locator(selector).InnerTextAsync();
        }

        public async Task WaitForElementAsync(string selector, int timeout = 5000)
        {
            await Page.Locator(selector).WaitForAsync(new LocatorWaitForOptions { Timeout = timeout });
        }

        public async Task<bool> IsElementVisibleAsync(string selector)
        {
            return await Page.Locator(selector).IsVisibleAsync();
        }
    }
}