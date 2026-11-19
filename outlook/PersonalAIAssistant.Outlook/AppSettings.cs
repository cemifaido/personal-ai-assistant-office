using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        /// <summary>Indirizzi modificati dall'utente per i provider compatibili OpenAI (assente = predefinito).</summary>
        public Dictionary<string, string> BaseUrls = new Dictionary<string, string>();
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

        /// <summary>
        /// Consensi a inviare la chiave di un provider a un indirizzo diverso da quello ufficiale:
        /// provider -> "origine|impronta della chiave". Cambiando server o chiave il consenso non vale piu'.
        /// </summary>
        public Dictionary<string, string> KeyApprovals = new Dictionary<string, string>();

        public static string OriginOf(string url)
        {
            Uri u;
            return Uri.TryCreate(url ?? "", UriKind.Absolute, out u) ? u.GetLeftPart(UriPartial.Authority).ToLowerInvariant() : "";
        }

        public static string ApprovalToken(string origin, string key)
        {
            using (var sha = SHA256.Create())
            {
                var h = sha.ComputeHash(Encoding.UTF8.GetBytes((key ?? "").Trim()));
                return origin + "|" + BitConverter.ToString(h, 0, 12).Replace("-", "");
            }
        }

        /// <summary>true se l'indirizzo e' quello ufficiale, il modello locale, o un indirizzo confermato per questa chiave.</summary>
        public static bool DestinationApproved(string provider, string url, string key, IDictionary<string, string> approvals)
        {
            if (!ModelCatalog.IsCompatible(provider) || provider == ModelCatalog.Local) return true;
            var origin = OriginOf(url);
            if (origin == OriginOf(ModelCatalog.CompatBaseUrls[provider])) return true;
            if (string.IsNullOrWhiteSpace(key)) return true;   // nessuna chiave da proteggere
            string token;
            return approvals.TryGetValue(provider, out token) && token == ApprovalToken(origin, key);
        }

        public bool IsDestinationApproved(string provider)
        {
            return DestinationApproved(provider, BaseUrlFor(provider), GetKey(provider), KeyApprovals);
        }

        /// <summary>Indirizzo da usare per un provider compatibile: quello salvato (se valido) o il predefinito.</summary>
        public string BaseUrlFor(string provider)
        {
            string url;
            if (BaseUrls.TryGetValue(provider, out url))
            {
                var ok = ModelCatalog.NormalizeBaseUrl(url, provider);
                if (ok != null) return ok;
            }
            string def;
            return ModelCatalog.CompatBaseUrls.TryGetValue(provider, out def) ? def : null;
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
            if (Provider == ModelCatalog.Local) return true;   // il modello locale non usa chiavi
            foreach (var p in ModelCatalog.ProviderNames.Keys) if (GetKey(p).Length > 0) return true;
            return false;
        }

        private static string BackupPath { get { return FilePath + ".bak"; } }
        private static string TempPath { get { return FilePath + ".tmp"; } }

        /// <summary>Se valorizzato, le impostazioni sono state recuperate da una copia (il file principale mancava o era danneggiato).</summary>
        public string RecoveredFrom { get; private set; }

        /// <summary>Legge un file di impostazioni; false se manca o non e' un JSON di impostazioni valido.</summary>
        private static bool TryReadInto(AppSettings s, string path)
        {
            try
            {
                if (!File.Exists(path)) return false;
                var data = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(path, Encoding.UTF8));
                if (data == null) return false;
                object v;
                if (data.TryGetValue("provider", out v) && v is string) s.Provider = (string)v;
                if (data.TryGetValue("models", out v) && v is Dictionary<string, object>)
                    foreach (var kv in (Dictionary<string, object>)v) s.Models[kv.Key] = Convert.ToString(kv.Value);
                if (data.TryGetValue("urls", out v) && v is Dictionary<string, object>)
                    foreach (var kv in (Dictionary<string, object>)v)
                        if (ModelCatalog.IsCompatible(kv.Key)) s.BaseUrls[kv.Key] = Convert.ToString(kv.Value);
                if (data.TryGetValue("keyApprovals", out v) && v is Dictionary<string, object>)
                    foreach (var kv in (Dictionary<string, object>)v) s.KeyApprovals[kv.Key] = Convert.ToString(kv.Value);
                if (data.TryGetValue("keys", out v) && v is Dictionary<string, object>)
                    foreach (var kv in (Dictionary<string, object>)v) s.protectedKeys[kv.Key] = Convert.ToString(kv.Value);
                return true;
            }
            catch { return false; }
        }

        public static AppSettings Load()
        {
            var s = new AppSettings();
            bool mainExists = File.Exists(FilePath);
            if (!TryReadInto(s, FilePath))
            {
                // Un file principale danneggiato si mette da parte (mai sovrascritto) per un recupero manuale.
                if (mainExists)
                {
                    try { File.Copy(FilePath, Path.Combine(FolderPath, "outlook-settings.danneggiato-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".json"), false); }
                    catch { }
                }
                // Copie di riserva: backup dell'ultimo salvataggio, poi il temporaneo di un salvataggio interrotto.
                s = new AppSettings();
                if (TryReadInto(s, BackupPath)) s.RecoveredFrom = "copia di backup";
                else
                {
                    s = new AppSettings();
                    if (TryReadInto(s, TempPath)) s.RecoveredFrom = "salvataggio interrotto";
                    else
                    {
                        s = new AppSettings();
                        if (!mainExists) { try { s.ImportFromLegacyMacro(); } catch { } }
                    }
                }
            }
            // A ogni caricamento: se restano valori della vecchia macro gia' importati, si ritenta la pulizia.
            // (Rilegge il file principale: se le impostazioni sono state recuperate da una copia non cancella nulla.)
            s.CleanLegacyMacro();
            if (!ModelCatalog.ProviderNames.ContainsKey(s.Provider)) s.Provider = "gemini";
            return s;
        }

        /// <summary>
        /// Salvataggio senza finestre di perdita: scrive un temporaneo, lo rilegge per verificarlo, poi sostituisce
        /// il file principale con File.Replace (che conserva la versione precedente come .bak). In nessun momento
        /// manca una copia valida delle impostazioni. Lancia un'eccezione se non riesce: il chiamante la mostra.
        /// </summary>
        public void Save()
        {
            Directory.CreateDirectory(FolderPath);
            var data = new Dictionary<string, object>
            {
                { "version", 1 },
                { "provider", Provider },
                { "models", Models },
                { "urls", BaseUrls },
                { "keyApprovals", KeyApprovals },
                { "keys", protectedKeys },
            };
            var bytes = new UTF8Encoding(false).GetBytes(new JavaScriptSerializer().Serialize(data));
            using (var fs = new FileStream(TempPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, FileOptions.WriteThrough))
            {
                fs.Write(bytes, 0, bytes.Length);
                fs.Flush(true);
            }
            if (!TryReadInto(new AppSettings(), TempPath))
                throw new IOException("Verifica del file delle impostazioni non riuscita: le impostazioni precedenti restano invariate.");
            if (File.Exists(FilePath)) File.Replace(TempPath, FilePath, BackupPath, true);
            else File.Move(TempPath, FilePath);
            RecoveredFrom = null;
        }

        /// <summary>
        /// Prima esecuzione: importa provider, modelli e chiavi della vecchia macro VBA
        /// (salvate in chiaro nel registro) e poi ne cancella SOLO i valori migrati.
        /// "Importate" e "rimosse dal registro" sono stati distinti: se la pulizia non riesce
        /// lo si dichiara e si ritenta a ogni caricamento.
        /// </summary>
        public bool ImportedFromMacro { get; private set; }

        /// <summary>true se nel registro restano ancora chiavi della vecchia macro in chiaro.</summary>
        public bool LegacyKeysRemain { get; private set; }

        private const string LegacyPath = @"Software\VB and VBA Program Settings\PersonalAIAssistant\Config";
        private const string LegacyRoot = @"Software\VB and VBA Program Settings\PersonalAIAssistant";

        private void ImportFromLegacyMacro()
        {
            using (var k = Registry.CurrentUser.OpenSubKey(LegacyPath, false))
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
            if (!ImportedFromMacro) return;
            // Prima si salva la copia cifrata, poi (in CleanLegacyMacro) si pulisce il registro.
            try { Save(); }
            catch { ImportedFromMacro = false; ImportSaveFailed = true; }
        }

        /// <summary>true se la copia cifrata delle chiavi importate non e' stata salvata: il registro resta intatto.</summary>
        public bool ImportSaveFailed { get; private set; }

        /// <summary>Chiavi in chiaro come risultano DAL FILE SALVATO (non dalla memoria); vuoto se il file manca o non si legge.</summary>
        private static Dictionary<string, string> ReadSavedKeys()
        {
            var result = new Dictionary<string, string>();
            try
            {
                if (!File.Exists(FilePath)) return result;
                var data = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(FilePath, Encoding.UTF8));
                object v;
                if (!data.TryGetValue("keys", out v) || !(v is Dictionary<string, object>)) return result;
                foreach (var kv in (Dictionary<string, object>)v)
                {
                    try
                    {
                        var plain = ProtectedData.Unprotect(Convert.FromBase64String(Convert.ToString(kv.Value)), Entropy, DataProtectionScope.CurrentUser);
                        result[kv.Key] = Encoding.UTF8.GetString(plain);
                    }
                    catch { /* chiave illeggibile: come se non ci fosse */ }
                }
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Rimuove dal registro i soli valori della vecchia macro (chiavi, provider, modelli) e verifica il risultato.
        /// Una chiave si cancella solo se, rileggendo il file salvato, vi si trova la STESSA chiave
        /// per lo stesso provider: se il salvataggio e' fallito o la chiave e' diversa, il registro resta intatto.
        /// </summary>
        private void CleanLegacyMacro()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(LegacyPath, true))
                {
                    if (k == null) { LegacyKeysRemain = false; return; }
                    bool saved = File.Exists(FilePath);
                    var savedKeys = ReadSavedKeys();
                    foreach (var p in ModelCatalog.ProviderNames.Keys)
                    {
                        var legacy = k.GetValue("key_" + p) as string;
                        string stored;
                        if (legacy != null && savedKeys.TryGetValue(p, out stored) && stored.Length > 0 &&
                            string.Equals(stored, legacy.Trim(), StringComparison.Ordinal))
                            k.DeleteValue("key_" + p, false);
                        if (saved) k.DeleteValue("model_" + p, false);
                    }
                    if (saved) k.DeleteValue("provider", false);
                }
                using (var k = Registry.CurrentUser.OpenSubKey(LegacyPath, false))
                {
                    LegacyKeysRemain = k != null && k.GetValueNames().Any(n => n.StartsWith("key_", StringComparison.OrdinalIgnoreCase));
                    if (k != null && k.ValueCount == 0 && k.SubKeyCount == 0)
                    {
                        k.Close();
                        Registry.CurrentUser.DeleteSubKey(LegacyPath, false);
                        using (var r = Registry.CurrentUser.OpenSubKey(LegacyRoot, false))
                            if (r != null && r.ValueCount == 0 && r.SubKeyCount == 0) { r.Close(); Registry.CurrentUser.DeleteSubKey(LegacyRoot, false); }
                    }
                }
            }
            catch
            {
                using (var k = Registry.CurrentUser.OpenSubKey(LegacyPath, false))
                    LegacyKeysRemain = k != null && k.GetValueNames().Any(n => n.StartsWith("key_", StringComparison.OrdinalIgnoreCase));
            }
        }
    }
}
