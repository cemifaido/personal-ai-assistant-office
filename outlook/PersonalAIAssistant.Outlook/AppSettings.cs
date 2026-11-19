using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.Win32;

namespace PersonalAIAssistant
{
    /// <summary>
    /// Impostazioni dell'add-in Outlook, per utente:
    /// %APPDATA%\PersonalAIAssistant\outlook-settings.json.
    /// Le chiavi API sono cifrate con DPAPI (ambito utente corrente): leggibili solo
    /// da questo utente Windows su questo PC. Non protegge da programmi eseguiti dallo stesso utente.
    /// </summary>
    internal sealed class AppSettings
    {
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("PersonalAIAssistant.Outlook.v1");

        public string Provider = "gemini";
        public Dictionary<string, string> Models = new Dictionary<string, string>();
        private Dictionary<string, string> protectedKeys = new Dictionary<string, string>();

        public static string FolderPath
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PersonalAIAssistant"); }
        }

        private static string FilePath { get { return Path.Combine(FolderPath, "outlook-settings.json"); } }

        public string ModelFor(string provider)
        {
            string id;
            if (Models.TryGetValue(provider, out id) && !string.IsNullOrEmpty(id)) return id;
            return ModelCatalog.DefaultFor(provider).Id;
        }

        public string GetKey(string provider)
        {
            string b64;
            if (!protectedKeys.TryGetValue(provider, out b64) || string.IsNullOrEmpty(b64)) return "";
            try
            {
                var plain = ProtectedData.Unprotect(Convert.FromBase64String(b64), Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plain);
            }
            catch { return ""; }
        }

        public void SetKey(string provider, string key)
        {
            if (string.IsNullOrWhiteSpace(key)) { protectedKeys.Remove(provider); return; }
            var enc = ProtectedData.Protect(Encoding.UTF8.GetBytes(key.Trim()), Entropy, DataProtectionScope.CurrentUser);
            protectedKeys[provider] = Convert.ToBase64String(enc);
        }

        public bool HasAnyKey()
        {
            foreach (var p in ModelCatalog.ProviderNames.Keys) if (GetKey(p).Length > 0) return true;
            return false;
        }

        public static AppSettings Load()
        {
            var s = new AppSettings();
            try
            {
                if (File.Exists(FilePath))
                {
                    var data = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(FilePath, Encoding.UTF8));
                    object v;
                    if (data.TryGetValue("provider", out v) && v is string) s.Provider = (string)v;
                    if (data.TryGetValue("models", out v) && v is Dictionary<string, object>)
                        foreach (var kv in (Dictionary<string, object>)v) s.Models[kv.Key] = Convert.ToString(kv.Value);
                    if (data.TryGetValue("keys", out v) && v is Dictionary<string, object>)
                        foreach (var kv in (Dictionary<string, object>)v) s.protectedKeys[kv.Key] = Convert.ToString(kv.Value);
                }
                else
                {
                    s.ImportFromLegacyMacro();
                }
            }
            catch { /* file danneggiato: si riparte dalle impostazioni predefinite */ }
            if (!ModelCatalog.ProviderNames.ContainsKey(s.Provider)) s.Provider = "gemini";
            return s;
        }

        public void Save()
        {
            Directory.CreateDirectory(FolderPath);
            var data = new Dictionary<string, object>
            {
                { "version", 1 },
                { "provider", Provider },
                { "models", Models },
                { "keys", protectedKeys },
            };
            var tmp = FilePath + ".tmp";
            File.WriteAllText(tmp, new JavaScriptSerializer().Serialize(data), Encoding.UTF8);
            if (File.Exists(FilePath)) File.Delete(FilePath);
            File.Move(tmp, FilePath);
        }

        /// <summary>
        /// Prima esecuzione: importa provider, modelli e chiavi della vecchia macro VBA
        /// (salvate in chiaro nel registro) e le cancella dal registro.
        /// </summary>
        public bool ImportedFromMacro { get; private set; }

        private void ImportFromLegacyMacro()
        {
            const string path = @"Software\VB and VBA Program Settings\PersonalAIAssistant\Config";
            using (var k = Registry.CurrentUser.OpenSubKey(path, true))
            {
                if (k == null) return;
                var prov = k.GetValue("provider") as string;
                if (!string.IsNullOrEmpty(prov)) Provider = prov;
                foreach (var p in ModelCatalog.ProviderNames.Keys)
                {
                    var m = k.GetValue("model_" + p) as string;
                    if (!string.IsNullOrEmpty(m)) Models[p] = m;
                    var key = k.GetValue("key_" + p) as string;
                    if (!string.IsNullOrEmpty(key)) { SetKey(p, key); ImportedFromMacro = true; }
                }
            }
            if (ImportedFromMacro)
            {
                Save();
                try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\VB and VBA Program Settings\PersonalAIAssistant", false); } catch { }
            }
        }
    }
}
