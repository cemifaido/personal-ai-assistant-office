# Eseguire con Windows PowerShell 5.1: compila i sorgenti reali in un ambiente isolato.
# Solo il percorso dati e il nome delle chiavi legacy vengono reindirizzati; nessun profilo reale viene usato.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$caseDir = Join-Path ([IO.Path]::GetTempPath()) ('paa-settings-test-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($caseDir) > $null
$source = [IO.File]::ReadAllText((Join-Path $root 'outlook/PersonalAIAssistant.Outlook/AppSettings.cs'))
$original = 'Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PersonalAIAssistant")'
if (-not $source.Contains($original)) { throw 'Percorso dati cambiato: aggiornare isolamento test.' }
$source = $source.Replace($original, '@"' + $caseDir.Replace('"', '""') + '"')
$source = $source.Replace('Software\VB and VBA Program Settings\PersonalAIAssistant', 'Software\PersonalAIAssistant.Tests\' + [Guid]::NewGuid().ToString('N'))
$catalog = [IO.File]::ReadAllText((Join-Path $root 'outlook/PersonalAIAssistant.Outlook/ModelCatalog.cs'))
$checks = @'
using System;
using System.IO;
using System.Reflection;
namespace PersonalAIAssistant {
    public static class SettingsChecks {
        static void Check(bool ok, string label) { if (!ok) throw new Exception(label); Console.WriteLine("PASS " + label); }
        public static void Run() {
            string path = Path.Combine(AppSettings.FolderPath, "outlook-settings.json");
            var s = new AppSettings();
            s.Models["openai"] = "first-model";
            s.Save();
            s.Models["openai"] = "second-model";
            s.Save();
            Check(File.Exists(path + ".bak") && !File.Exists(path + ".tmp"), "atomic save creates backup, consumes temporary");
            Check(AppSettings.Load().Models["openai"] == "second-model", "main settings take priority");
            var read = typeof(AppSettings).GetMethod("TryReadInto", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            string probe = Path.Combine(AppSettings.FolderPath, "probe.json");
            foreach (string bad in new[] { "{}", "null", "{broken", "{\"version\":1,\"provider\":\"gemini\",\"models\":{},\"keys\":[]}", "{\"version\":2,\"provider\":\"gemini\",\"models\":{},\"keys\":{}}" }) {
                File.WriteAllText(probe, bad);
                Check(!(bool)read.Invoke(null, new object[] {new AppSettings(), probe}), "invalid settings rejected: " + bad);
            }
            File.WriteAllText(probe, "{\"version\":1,\"provider\":\"gemini\",\"models\":{},\"keys\":{}}");
            Check((bool)read.Invoke(null, new object[] {new AppSettings(), probe}), "older schema without urls/approvals accepted");
            File.WriteAllText(path, "{}");
            var restored = AppSettings.Load();
            Check(restored.RecoveredFrom != null && restored.Models["openai"] == "first-model", "invalid main recovers backup");
            Check(Directory.GetFiles(AppSettings.FolderPath, "outlook-settings.danneggiato-*.json").Length > 0, "damaged file preserved");
            restored.Save();
            Check(restored.RecoveredFrom == null && AppSettings.Load().Models["openai"] == "first-model", "recovered settings can be saved");
            File.Move(path, path + ".tmp");
            File.Delete(path + ".bak");
            Check(AppSettings.Load().RecoveredFrom == "salvataggio interrotto", "temporary recovered when main and backup absent");
            Console.WriteLine("All Outlook settings checks passed");
        }
    }
}
'@
Add-Type -TypeDefinition ($source + "`n" + $catalog.Replace('using System;', '').Replace('using System.Collections.Generic;', '').Replace('using System.Linq;', '') + "`n" + $checks.Replace('using System;', '').Replace('using System.IO;', '').Replace('using System.Reflection;', '')) -ReferencedAssemblies System.Web.Extensions, System.Security, System.Core
[PersonalAIAssistant.SettingsChecks]::Run()
Write-Host "Artefatti isolati: $caseDir"
