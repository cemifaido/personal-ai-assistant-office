using System.Collections.Generic;
using System.Linq;

namespace PersonalAIAssistant
{
    /// <summary>
    /// Catalogo dei modelli. Tenerlo allineato a AI_CATALOG in public/ai.js.
    /// </summary>
    internal static class ModelCatalog
    {
        internal sealed class Model
        {
            public string Provider;
            public string Id;
            public string Label;
            public override string ToString() { return Label; }
        }

        public static readonly Dictionary<string, string> ProviderNames = new Dictionary<string, string>
        {
            { "gemini", "Google Gemini" },
            { "openai", "OpenAI" },
            { "claude", "Anthropic Claude" },
        };

        public static readonly Dictionary<string, string> KeyUrls = new Dictionary<string, string>
        {
            { "gemini", "https://aistudio.google.com/apikey" },
            { "openai", "https://platform.openai.com/api-keys" },
            { "claude", "https://console.anthropic.com/settings/keys" },
        };

        public static readonly List<Model> Models = new List<Model>
        {
            new Model { Provider = "gemini", Id = "gemini-3.1-flash-lite", Label = "Gemini 3.1 Flash Lite (veloce)" },
            new Model { Provider = "gemini", Id = "gemini-3.8-flash", Label = "Gemini 3.8 Flash" },
            new Model { Provider = "gemini", Id = "gemini-2.5-pro", Label = "Gemini 2.5 Pro" },
            new Model { Provider = "openai", Id = "gpt-4o-mini", Label = "GPT-4o Mini" },
            new Model { Provider = "openai", Id = "gpt-4o", Label = "GPT-4o" },
            new Model { Provider = "openai", Id = "o3-mini", Label = "o3-mini (ragionamento)" },
            new Model { Provider = "claude", Id = "claude-sonnet-5-5", Label = "Claude Sonnet 5.5 (scrittura)" },
            new Model { Provider = "claude", Id = "claude-haiku-5-5", Label = "Claude Haiku 5.5 (rapido)" },
            new Model { Provider = "claude", Id = "claude-opus-5-5", Label = "Claude Opus 5.5 (massima qualita')" },
        };

        public static Model Find(string provider, string id)
        {
            return Models.FirstOrDefault(m => m.Provider == provider && m.Id == id);
        }

        public static Model DefaultFor(string provider)
        {
            return Models.First(m => m.Provider == provider);
        }
    }
}
