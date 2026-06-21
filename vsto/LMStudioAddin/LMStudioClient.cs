using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace LMStudioAddin
{
    public class LMStudioModel
    {
        public string Id { get; set; }
        public string Object { get; set; }
    }

    public class ChatMessage
    {
        public string Role { get; set; }
        public string Content { get; set; }
    }

    public class ChatRequest
    {
        public string Model { get; set; }
        public List<ChatMessage> Messages { get; set; }
        public double Temperature { get; set; } = 0.7;
        public bool Stream { get; set; } = false;
    }

    /// <summary>
    /// Thin wrapper around LMStudio's OpenAI-compatible local REST API.
    /// Base URL defaults to http://localhost:1234/v1.
    /// </summary>
    public class LMStudioClient : IDisposable
    {
        private readonly HttpClient _http;
        private string _baseUrl;

        public string BaseUrl
        {
            get => _baseUrl;
            set => _baseUrl = value.TrimEnd('/');
        }

        public LMStudioClient(string baseUrl = "http://localhost:1234/v1")
        {
            _baseUrl = baseUrl.TrimEnd('/');
            _http = new HttpClient { Timeout = TimeSpan.FromSeconds(120) };
            _http.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
        }

        /// <summary>Returns all models currently loaded in LMStudio.</summary>
        public async Task<List<LMStudioModel>> GetModelsAsync()
        {
            var response = await _http.GetAsync($"{_baseUrl}/models");
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            var models = new List<LMStudioModel>();
            if (doc.RootElement.TryGetProperty("data", out var data))
            {
                foreach (var item in data.EnumerateArray())
                {
                    models.Add(new LMStudioModel
                    {
                        Id = item.GetProperty("id").GetString(),
                        Object = item.TryGetProperty("object", out var o) ? o.GetString() : ""
                    });
                }
            }
            return models;
        }

        /// <summary>Single-shot chat completion (non-streaming).</summary>
        public async Task<string> ChatAsync(ChatRequest request,
            CancellationToken cancellationToken = default)
        {
            request.Stream = false;
            var body = SerializeRequest(request);
            var content = new StringContent(body, Encoding.UTF8, "application/json");
            var response = await _http.PostAsync($"{_baseUrl}/chat/completions", content, cancellationToken);
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(json);
            return doc.RootElement
                       .GetProperty("choices")[0]
                       .GetProperty("message")
                       .GetProperty("content")
                       .GetString() ?? string.Empty;
        }

        /// <summary>
        /// Streaming chat completion — invokes <paramref name="onChunk"/> for each
        /// text token as it arrives, then returns the full assembled response.
        /// </summary>
        public async Task<string> ChatStreamAsync(
            ChatRequest request,
            Action<string> onChunk,
            CancellationToken cancellationToken = default)
        {
            request.Stream = true;
            var body = SerializeRequest(request);
            var requestMessage = new HttpRequestMessage(HttpMethod.Post,
                $"{_baseUrl}/chat/completions")
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json")
            };
            requestMessage.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("text/event-stream"));

            using var response = await _http.SendAsync(requestMessage,
                HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            response.EnsureSuccessStatusCode();

            using var stream = await response.Content.ReadAsStreamAsync();
            using var reader = new System.IO.StreamReader(stream);

            var fullText = new StringBuilder();
            string line;
            while ((line = await reader.ReadLineAsync()) != null)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrWhiteSpace(line) || line == "data: [DONE]") continue;
                if (line.StartsWith("data: "))
                {
                    try
                    {
                        using var doc = JsonDocument.Parse(line.Substring(6));
                        if (doc.RootElement
                               .GetProperty("choices")[0]
                               .GetProperty("delta")
                               .TryGetProperty("content", out var c))
                        {
                            var chunk = c.GetString();
                            if (!string.IsNullOrEmpty(chunk))
                            {
                                fullText.Append(chunk);
                                onChunk?.Invoke(chunk);
                            }
                        }
                    }
                    catch { /* skip malformed SSE chunks */ }
                }
            }
            return fullText.ToString();
        }

        private static string SerializeRequest(ChatRequest req)
        {
            var messages = new System.Text.StringBuilder();
            foreach (var m in req.Messages)
            {
                if (messages.Length > 0) messages.Append(",");
                messages.Append($"{{"role":"{m.Role}","content":{JsonEncodedText.Encode(m.Content)}}}");
            }
            return $"{{"model":"{req.Model}","messages":[{messages}]," +
                   $""temperature":{req.Temperature},"stream":{req.Stream.ToString().ToLower()}}}";
        }

        public void Dispose() => _http?.Dispose();
    }
}
