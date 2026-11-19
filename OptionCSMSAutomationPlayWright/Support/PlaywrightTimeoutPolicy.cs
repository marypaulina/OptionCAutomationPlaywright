using Microsoft.Playwright;

namespace OptionCSMSAutomationPlayWright.Support
{
    internal static class PlaywrightTimeoutPolicy
    {
        public const int ActionTimeoutMs = 60000;
        public const int NavigationTimeoutMs = 120000;

        public static void Configure(IBrowserContext context)
        {
            context.SetDefaultTimeout(ActionTimeoutMs);
            context.SetDefaultNavigationTimeout(NavigationTimeoutMs);
        }
    }
}
