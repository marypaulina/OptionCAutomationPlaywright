using NUnit.Framework;
using OptionCSMSAutomationPlayWright.Helpers;
using OptionCSMSAutomationPlayWright.Helpers;

namespace OptionCSMSAutomationPlayWright.Tests
{
    public class AITest
    {
        [Test]
        public async Task Verify_AI_Response()
        {
            AIHelper ai = new AIHelper();

            string response = await ai.AskAI(
                "Generate 3 invalid login test cases");

            Console.WriteLine(response);
        }
    }
}