using AventStack.ExtentReports;
using AventStack.ExtentReports.Gherkin.Model;
using AventStack.ExtentReports.Reporter;
using BoDi;
using Microsoft.Playwright;
using System.Diagnostics;
using TechTalk.SpecFlow;

namespace OptionCSMSAutomationPlayWright.Hooks
{
    [Binding]
    public sealed class SpecflowSeleniumHooks
    {
        private static ExtentTest featureName;
        public static ExtentTest scenario;
        private static AventStack.ExtentReports.ExtentReports extent;
        private static string reportPath = System.IO.Directory.GetParent(@"../../../").FullName + Path.DirectorySeparatorChar + "Result" + Path.DirectorySeparatorChar + "Result_" + DateTime.Now.ToString("ddMMyyyyHHmmss");

        private readonly IObjectContainer container;
        private IPlaywright _playwright;
        private IBrowser _browser;
        private IPage _page;

        private const string ScreenshotsDirectory = "Screenshots";
        private const string ScreenshotFileExtension = ".png";
        private const string ChromeExecutablePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe";

        public SpecflowSeleniumHooks(IObjectContainer container)
        {
            this.container = container;
        }

        [BeforeTestRun]
        public static void BeforeTestRun()
        {
            ExtentHtmlReporter htmlReport = new ExtentHtmlReporter(reportPath);
            extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlReport);
        }

        [BeforeFeature]
        public static void BeforeFeature(FeatureContext featureContext)
        {
            // Create dynamic feature name
            featureName = extent.CreateTest<Feature>(featureContext.FeatureInfo.Title);
        }

        [BeforeScenario]
        public async Task CreatePlaywrightBrowser(ScenarioContext scenarioContext)
        {
            Console.WriteLine("BeforeScenario");
            scenario = featureName.CreateNode<Scenario>(scenarioContext.ScenarioInfo.Title);

            // Initialize Playwright
            _playwright = await Playwright.CreateAsync();

            // Launch the browser
            _browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = false, // Set to true for headless mode if needed
                ExecutablePath = ChromeExecutablePath

            });

            // Create a new context
            var context = await _browser.NewContextAsync();

            // Create a new page within the context
            _page = await context.NewPageAsync();

            // Set a large viewport size
            await _page.SetViewportSizeAsync(1500, 700); // Set according to your screen size

            // Optionally, you can reposition the browser window using the DevTools Protocol
            // Note: This doesn't directly maximize the window but allows for custom sizing
            await _page.EvaluateAsync(@"window.resizeTo(screen.width, screen.height);");

            // Bring the page to the front for visibility
            await _page.BringToFrontAsync();

            // Register the page instance
            container.RegisterInstanceAs<IPage>(_page);
        }

        [AfterStep]
        public async Task InsertReportingSteps(ScenarioContext scenarioContext)
         {
            var stepType = ScenarioStepContext.Current.StepInfo.StepDefinitionType.ToString();
            if (scenarioContext.TestError == null)
            {
                switch (stepType)
                {
                    case "Given":
                        scenario.CreateNode<Given>(ScenarioStepContext.Current.StepInfo.Text);
                        break;
                    case "When":
                        scenario.CreateNode<When>(ScenarioStepContext.Current.StepInfo.Text);
                        break;
                    case "Then":
                        scenario.CreateNode<Then>(ScenarioStepContext.Current.StepInfo.Text);
                        break;
                    case "And":
                        scenario.CreateNode<And>(ScenarioStepContext.Current.StepInfo.Text);
                        break;
                }
            }
            else
            {
                // Capture screenshot on error
                try
                {
                    var screenshotBytes = await _page.ScreenshotAsync(new PageScreenshotOptions { FullPage = true });
                    var screenshotPath = Path.Combine(ScreenshotsDirectory, $"{scenarioContext.ScenarioInfo.Title}_{DateTime.Now:yyyyMMdd_HHmmss}{ScreenshotFileExtension}");

                    if (!Directory.Exists(ScreenshotsDirectory))
                        Directory.CreateDirectory(ScreenshotsDirectory);

                    await File.WriteAllBytesAsync(screenshotPath, screenshotBytes);
                    Console.WriteLine($"Screenshot saved to: {screenshotPath}");

                    // Add failure details to Extent Report
                    switch (stepType)
                    {
                        case "Given":
                            scenario.CreateNode<Given>(ScenarioStepContext.Current.StepInfo.Text).Fail(scenarioContext.TestError.Message);
                            break;
                        case "When":
                            scenario.CreateNode<When>(ScenarioStepContext.Current.StepInfo.Text).Fail(scenarioContext.TestError.Message);
                            break;
                        case "Then":
                            scenario.CreateNode<Then>(ScenarioStepContext.Current.StepInfo.Text).Fail(scenarioContext.TestError.Message);
                            break;
                        case "And":
                            scenario.CreateNode<And>(ScenarioStepContext.Current.StepInfo.Text).Fail(scenarioContext.TestError.Message);
                            break;
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to capture screenshot: {e.Message}");
                }
            }
        }

        [AfterScenario]
        public async Task DestroyPlaywrightBrowser()
        {
            if (_page != null)
            {
                try
                {
                    await _page.CloseAsync();
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to close page: {e.Message}");
                }
                _page = null; // Avoid repeated disposal
            }

            if (_browser != null)
            {
                try
                {
                    await _browser.CloseAsync();
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to close browser: {e.Message}");
                }
                _browser = null; // Avoid repeated disposal
            }

            if (_playwright != null)
            {
                try
                {
                    _playwright.Dispose();
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to dispose Playwright: {e.Message}");
                }
                _playwright = null; // Avoid repeated disposal
            }
        }

        [AfterFeature]
        public static void AfterFeature()
        {
            extent.Flush();
        }

        [AfterTestRun]
        public static void AfterTestRun()
        {
            Process[] playwrightProcesses = Process.GetProcessesByName("playwright");
            foreach (var playwrightProcess in playwrightProcesses)
            {
                playwrightProcess.Kill();
            }
        }
    }
}
