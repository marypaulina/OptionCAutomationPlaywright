using Microsoft.Playwright;

namespace OptionCSMSAutomationPlayWright.SISPages
{
    public static class OptionCWebElement
    {
        public static async Task<IElementHandle> GetElementByIdAsync(IPage page, string controlId)
        {
            await page.WaitForSelectorAsync($"//*[@id='{controlId}']");
            return await page.QuerySelectorAsync($"//*[@id='{controlId}']");
        }

        public static async Task<IElementHandle> GetElementByNameAsync(IPage page, string controlName)
        {
            await page.WaitForSelectorAsync($"//*[@name='{controlName}']");
            return await page.QuerySelectorAsync($"//*[@name='{controlName}']");
        }

        public static async Task<bool> IsElementPresentAsync(IPage page, string elementId)
        {
            var element = await page.QuerySelectorAsync($"#{elementId}");
            return element != null;
        }

        public static async Task<bool> IsElementPresentByXpathAsync(IPage page, string xpath)
        {
            var element = await page.QuerySelectorAsync(xpath);
            return element != null;
        }

        public class LoginWebElement
        {
            public static async Task<IElementHandle> GetUsernameFieldAsync(IPage page)
            {
                return await GetElementByIdAsync(page, "username");
            }

            public static async Task<IElementHandle> GetPasswordFieldAsync(IPage page)
            {
                return await GetElementByIdAsync(page, "password");
            }

            public static async Task<IElementHandle> GetLoginButtonAsync(IPage page)
            {
                return await GetElementByIdAsync(page, "Login");
            }
        }
    }
}