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
        /*private IPlaywright _playwright;
        private IBrowser _browser;
        private IPage _page;*/

        private IPlaywright _playwright;
        private IBrowser _browser;
        private IBrowserContext _context;  // Added
        private IPage _page;

        private const string ScreenshotsDirectory = "Screenshots";
        private const string ScreenshotFileExtension = ".png";
        private const string ChromeExecutablePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe";
        private readonly ScenarioContext _scenarioContext;

        //public SpecflowSeleniumHooks(IObjectContainer container, ScenarioContext scenarioContext)
        //{
        //    this.container = container;
        //    _scenarioContext = scenarioContext;
        //}
        public SpecflowSeleniumHooks(IObjectContainer container)
        {
            this.container = container;
            //_scenarioContext = scenarioContext;
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

            if (_playwright == null)  // Ensure only one instance is created
                _playwright = await Playwright.CreateAsync();

            if (_browser == null)
            {
                _browser = await _playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                {
                    Headless = false,
                    ExecutablePath = ChromeExecutablePath
                });
            }

            _context = await _browser.NewContextAsync();
            _page = await _context.NewPageAsync();

            await _page.SetViewportSizeAsync(1500, 700);
            await _page.BringToFrontAsync();

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
                _page = null;
            }

            if (_browser != null)
            {
                try
                {
                    // Commenting this out for debugging
                    // await _browser.CloseAsync();
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to close browser: {e.Message}");
                }
                _browser = null;
            }

            if (_playwright != null)
            {
                try
                {
                    // Commenting this out for debugging
                    // _playwright.Dispose();
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Failed to dispose Playwright: {e.Message}");
                }
                _playwright = null;
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
