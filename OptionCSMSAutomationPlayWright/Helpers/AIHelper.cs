using Microsoft.Extensions.Configuration;
using OpenAI.Chat;

namespace OptionCSMSAutomationPlayWright.Helpers
{
    public class AIHelper
    {
        private readonly string _apiKey;

        public AIHelper()
        {
            var config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();

            _apiKey = config["OpenAI:ApiKey"];
        }

        public async Task<string> AskAI(string prompt)
        {
            var client = new ChatClient(
                model: "gpt-4o-mini",
                apiKey: _apiKey);

            var response = await client.CompleteChatAsync(prompt);

            return response.Value.Content[0].Text;
        }
    }
}