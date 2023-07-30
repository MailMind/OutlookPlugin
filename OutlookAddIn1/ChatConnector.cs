using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using Newtonsoft.Json;
using System.Net;

namespace OutlookAddIn1
{
    public class ChatConnector
    {
        private static readonly HttpClient httpClient = new HttpClient();


        public void Connect()
        {
        }

        public async Task<string> GetChatResponse(string userInput)
        {
            var api = new OpenAI_API.OpenAIAPI("");
            string result = await api.Completions.GetCompletion("Tell me a Joke");
            return result;

            // Set up the API endpoint
            //string endpoint = "https://api.openai.com/v1/chat/completions";

            //// Set up the request payload
            //string requestBody = $"{{\"model\": \"gpt-3.5-turbo\", \"messages\": [{{\"role\": \"system\", \"content\": \"You are a helpful assistant.\"}}, {{\"role\": \"user\", \"content\": \"{userInput}\"}}]}}";

            //// Set up the HTTP request
            //StringContent content = new StringContent(requestBody, Encoding.UTF8, "application/json");

            //// Send the API request
            //HttpResponseMessage response = await httpClient.PostAsync(endpoint, content);

            //// Read the response content
            //string responseContent = await response.Content.ReadAsStringAsync();

            //// Extract and return the generated message from the response
            //dynamic responseObject = Newtonsoft.Json.JsonConvert.DeserializeObject(responseContent);
            //string generatedMessage = responseObject.choices[0].message.content;
            //return generatedMessage;
        }
    }
}
