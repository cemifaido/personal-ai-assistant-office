using System;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace PaiOutlookSpike
{
    /// <summary>
    /// Spike: scheda "Personal AI (prova)" nella finestra principale di Outlook.
    /// Serve solo a verificare installazione, caricamento, trust e rete.
    /// </summary>
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private const string RibbonXml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='PaiSpikeTab' label='Personal AI (prova)'>
        <group id='PaiSpikeGroup' label='Prova tecnica'>
          <button id='PaiHello' label='Ciao' size='large' imageMso='Help' onAction='OnHello' />
          <button id='PaiNet' label='Prova rete' size='large' imageMso='WebPagePreview' onAction='OnNetTest' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public string GetCustomUI(string ribbonID)
        {
            // Solo la finestra principale (Explorer); nessuna modifica alle finestre dei messaggi.
            return ribbonID == "Microsoft.Outlook.Explorer" ? RibbonXml : null;
        }

        public void OnHello(Office.IRibbonControl control)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            MessageBox.Show(
                "Personal AI Assistant - prova tecnica\n\n" +
                "Versione add-in: " + version + "\n" +
                "Processo: " + (Environment.Is64BitProcess ? "64 bit" : "32 bit") + "\n" +
                "Utente: " + Environment.UserName + "\n" +
                ".NET CLR: " + Environment.Version,
                "Personal AI (prova)", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public async void OnNetTest(Office.IRibbonControl control)
        {
            // Chiamata asincrona: non blocca Outlook. 401 senza chiave = rete e TLS funzionano.
            string result;
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                using (var http = new HttpClient { Timeout = TimeSpan.FromSeconds(20) })
                {
                    var r1 = await Probe(http, "https://api.anthropic.com/v1/models");
                    var r2 = await Probe(http, "https://generativelanguage.googleapis.com/v1beta/models");
                    var r3 = await Probe(http, "https://api.openai.com/v1/models");
                    result = "Anthropic: " + r1 + "\nGoogle Gemini: " + r2 + "\nOpenAI: " + r3 +
                             "\n\n(401/403 senza chiave = connessione HTTPS riuscita)";
                }
            }
            catch (Exception ex)
            {
                result = "Errore: " + ex.Message;
            }
            MessageBox.Show(result, "Personal AI (prova) - rete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static async Task<string> Probe(HttpClient http, string url)
        {
            try
            {
                using (var resp = await http.GetAsync(url).ConfigureAwait(false))
                    return "HTTP " + (int)resp.StatusCode;
            }
            catch (Exception ex)
            {
                return "ERRORE " + (ex.InnerException?.Message ?? ex.Message);
            }
        }
    }
}
