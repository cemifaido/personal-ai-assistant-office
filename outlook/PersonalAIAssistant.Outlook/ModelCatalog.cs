using System;
using System.Collections.Generic;
using System.Linq;

namespace PersonalAIAssistant
{
    /// <summary>
    /// Catalogo dei provider e dei modelli. Tenerlo allineato a AI_CATALOG / AI_COMPAT in public/ai.js.
    /// Nelle Impostazioni modello e indirizzo dei provider compatibili OpenAI sono modificabili dall'utente:
    /// se un provider cambia nomi o indirizzi, si puo' correggere senza aspettare una nuova versione.
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
            { "qwen", "Qwen (Alibaba)" },
            { "minimax", "MiniMax" },
            { "deepseek", "DeepSeek" },
            { "openrouter", "OpenRouter" },
            { "locale", "Modello locale o in rete interna (Ollama, llama-server, LM Studio)" },
        };

        /// <summary>Server sul PC o nella rete interna: chiave facoltativa, indirizzo solo locale o privato.</summary>
        public const string Local = "locale";

        public static readonly Dictionary<string, string> KeyUrls = new Dictionary<string, string>
        {
            { "gemini", "https://aistudio.google.com/apikey" },
            { "openai", "https://platform.openai.com/api-keys" },
            { "claude", "https://console.anthropic.com/settings/keys" },
            { "qwen", "https://www.alibabacloud.com/help/en/model-studio/get-api-key" },
            { "minimax", "https://platform.minimax.io/user-center/basic-information/interface-key" },
            { "deepseek", "https://platform.deepseek.com/api_keys" },
            { "openrouter", "https://openrouter.ai/settings/keys" },
            { "locale", "https://ollama.com/download" },
        };

        /// <summary>Provider con API compatibile OpenAI: indirizzo predefinito (modificabile nelle Impostazioni).</summary>
        public static readonly Dictionary<string, string> CompatBaseUrls = new Dictionary<string, string>
        {
            { "qwen", "https://dashscope-intl.aliyuncs.com/compatible-mode/v1" },
            { "minimax", "https://api.minimax.io/v1" },
            { "deepseek", "https://api.deepseek.com" },
            { "openrouter", "https://openrouter.ai/api/v1" },
            { "locale", "http://localhost:11434/v1" },   // Ollama; llama-server :8080/v1, LM Studio :1234/v1
        };

        /// <summary>Provider con sede in Cina: si chiede conferma prima di salvarne la chiave.</summary>
        public static readonly Dictionary<string, string> ChinaProviders = new Dictionary<string, string>
        {
            { "qwen", "Qwen (Alibaba Cloud): servizio internazionale, con server a Singapore" },
            { "minimax", "MiniMax: azienda con sede a Shanghai" },
            { "deepseek", "DeepSeek: azienda con sede a Hangzhou, server in Cina" },
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
            new Model { Provider = "qwen", Id = "qwen-plus", Label = "Qwen Plus" },
            new Model { Provider = "qwen", Id = "qwen3-max", Label = "Qwen3 Max" },
            new Model { Provider = "qwen", Id = "qwen-turbo", Label = "Qwen Turbo (rapido)" },
            new Model { Provider = "minimax", Id = "MiniMax-M3", Label = "MiniMax M3" },
            new Model { Provider = "minimax", Id = "MiniMax-M2.7", Label = "MiniMax M2.7" },
            new Model { Provider = "deepseek", Id = "deepseek-v4-flash", Label = "DeepSeek V4 Flash" },
            new Model { Provider = "deepseek", Id = "deepseek-v4-pro", Label = "DeepSeek V4 Pro" },
            new Model { Provider = "openrouter", Id = "openrouter/auto", Label = "OpenRouter: scelta automatica" },
            new Model { Provider = "locale", Id = "qwen3:8b", Label = "qwen3:8b (esempio: scrivi il nome del tuo modello)" },
        };

        /// <summary>Questo PC, indirizzo privato (10/8, 172.16/12, 192.168/16, IPv6 fc00::/7 e fe80::/10) o nome senza dominio.</summary>
        public static bool IsInternalHost(Uri u)
        {
            if (u.IsLoopback) return true;
            System.Net.IPAddress ip;
            if (System.Net.IPAddress.TryParse(u.IdnHost.Trim('[', ']'), out ip))
            {
                var b = ip.GetAddressBytes();
                if (b.Length == 4)
                    return b[0] == 10 || (b[0] == 172 && b[1] >= 16 && b[1] <= 31) || (b[0] == 192 && b[1] == 168);
                return (b[0] & 0xFE) == 0xFC || (b[0] == 0xFE && (b[1] & 0xC0) == 0x80);
            }
            return u.HostNameType == UriHostNameType.Dns && u.Host.IndexOf('.') < 0;   // es. http://server-ai:11434
        }

        public static bool IsCompatible(string provider) { return CompatBaseUrls.ContainsKey(provider); }

        public static Model Find(string provider, string id)
        {
            return Models.FirstOrDefault(m => m.Provider == provider && m.Id == id);
        }

        public static Model DefaultFor(string provider)
        {
            return Models.First(m => m.Provider == provider);
        }

        /// <summary>
        /// Indirizzo accettabile per un provider compatibile: https, oppure http solo verso questo PC
        /// (localhost / 127.0.0.1 / ::1, per un modello locale). Restituisce l'indirizzo normalizzato o null.
        /// Toglie "/chat/completions" o "/models" finali se l'utente ha incollato l'indirizzo completo.
        /// </summary>
        public static string NormalizeBaseUrl(string input) { return NormalizeBaseUrl(input, null); }

        /// <summary>
        /// Provider su Internet: solo https. Modello locale: http o https, ma solo verso questo PC o la rete interna
        /// (indirizzi privati IPv4/IPv6 o nome di computer senza dominio): le chiavi e i testi non escono dalla rete.
        /// </summary>
        public static string NormalizeBaseUrl(string input, string provider)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;
            Uri u;
            if (!Uri.TryCreate(input.Trim(), UriKind.Absolute, out u)) return null;
            if (provider == Local)
            {
                if (!IsInternalHost(u)) return null;
                if (u.Scheme != Uri.UriSchemeHttp && u.Scheme != Uri.UriSchemeHttps) return null;
            }
            else if (u.Scheme != Uri.UriSchemeHttps) return null;
            if (!string.IsNullOrEmpty(u.UserInfo) || !string.IsNullOrEmpty(u.Query) || !string.IsNullOrEmpty(u.Fragment)) return null;
            var s = u.GetLeftPart(UriPartial.Path).TrimEnd('/');
            foreach (var tail in new[] { "/chat/completions", "/models" })
                if (s.EndsWith(tail, StringComparison.OrdinalIgnoreCase)) s = s.Substring(0, s.Length - tail.Length);
            return s;
        }
    }
}
