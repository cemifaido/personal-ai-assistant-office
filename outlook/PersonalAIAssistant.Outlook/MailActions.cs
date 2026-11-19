using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PersonalAIAssistant
{
    /// <summary>
    /// Azioni sulla posta. Il modello a oggetti di Outlook si usa SOLO sul thread dell'interfaccia:
    /// prima si leggono i dati, poi si attende la chiamata HTTP (asincrona), poi si crea la bozza.
    /// Nessun invio automatico: si apre sempre una bozza da rileggere.
    /// </summary>
    internal static class MailActions
    {
        private const int MaxCharsPerMail = 15000;
        private const int MaxCharsTotal = 80000;
        private const string Title = "Personal AI Assistant";

        private const string SystemPrompt =
            "Sei un assistente per la posta elettronica integrato in Outlook. Rispondi in italiano, in modo chiaro e professionale. " +
            "Puoi usare **grassetto** ed elenchi puntati con '- '. Non inventare date, scadenze o impegni assenti dalle email.";

        private static bool busy;

        // ---------------- punti di ingresso dai pulsanti ----------------
        // Gli unici "async void" sono questi: avvolgono TUTTA l'azione, compreso l'uso del modello
        // a oggetti dopo l'attesa (email eliminata/spostata, Outlook in chiusura...), cosi' nessuna
        // eccezione arriva al ciclo dei messaggi di Outlook.

        public static async void SummarizeSelection(Outlook.Application app)
        {
            await Guard(() => SummarizeCoreAsync(app));
        }

        public static async void Reply(Outlook.Application app, string intent)
        {
            await Guard(() => ReplyCoreAsync(app, intent));
        }

        private static async Task Guard(Func<Task> action)
        {
            try
            {
                await action();
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                SafeMessage("L'email non e' piu' disponibile (spostata, eliminata o cartella chiusa durante l'elaborazione). Riprova.", MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                SafeMessage("Errore imprevisto: " + ex.Message, MessageBoxIcon.Error);
            }
            finally
            {
                busy = false;
            }
        }

        private static void SafeMessage(string text, MessageBoxIcon icon)
        {
            try { MessageBox.Show(text, Title, MessageBoxButtons.OK, icon); } catch { /* Outlook in chiusura */ }
        }

        // ---------------- riassunto ----------------

        private static async Task SummarizeCoreAsync(Outlook.Application app)
        {
            if (busy) return;
            var exp = app.ActiveExplorer();
            var mails = SelectedMails(exp);
            if (mails.Count == 0)
            {
                MessageBox.Show("Seleziona una o piu' email (Ctrl+clic o Maiusc+clic).", Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            mails = mails.OrderBy(m => m.ReceivedTime).ToList();

            var sb = new StringBuilder();
            int truncated = 0, excluded = 0;
            for (int i = 0; i < mails.Count; i++)
            {
                var body = mails[i].Body ?? "";
                if (body.Length > MaxCharsPerMail) { body = body.Substring(0, MaxCharsPerMail) + "\n[... email troncata ...]"; truncated++; }
                if (sb.Length + body.Length > MaxCharsTotal) { excluded = mails.Count - i; break; }
                sb.Append("=== EMAIL ").Append(i + 1).Append(" di ").Append(mails.Count).Append('\n')
                  .Append("Da: ").Append(mails[i].SenderName).Append('\n')
                  .Append("Data: ").Append(mails[i].ReceivedTime.ToString("dd/MM/yyyy HH:mm")).Append('\n')
                  .Append("Oggetto: ").Append(mails[i].Subject).Append("\n\n").Append(body).Append("\n\n");
            }
            var warnings = (truncated > 0 ? truncated + " email troncate per lunghezza. " : "") +
                           (excluded > 0 ? excluded + " email piu' recenti escluse per limite di lunghezza." : "");
            var subject = mails[0].Subject;
            int count = mails.Count;

            var prompt = "Sono state fornite " + count + " email di una discussione, in ordine cronologico. Scrivi un riassunto esecutivo con queste sezioni:\n" +
                         "1. Tema e obiettivo principale\n" +
                         "2. Cronologia dei punti chiave (indica chi ha scritto cosa)\n" +
                         "3. Decisioni prese, azioni richieste e scadenze (solo se presenti; altrimenti scrivi che non ce ne sono)";

            var result = await RunAsync("Riassunto di " + count + " email...", prompt, sb.ToString());
            if (result == null) return;

            var draft = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);
            draft.Subject = "[Sintesi AI - " + count + " email] " + subject;
            draft.HTMLBody =
                "<div style=\"font-family:Aptos,Calibri,sans-serif;font-size:11pt;color:#000;line-height:1.5\">" +
                "<div style=\"background:#f1f3f4;padding:12px;border-left:4px solid #1a73e8;margin-bottom:16px\">" +
                "<b style=\"color:#1a73e8\">SINTESI ESECUTIVA (" + count + " email analizzate)</b><br>" +
                "<span style=\"font-size:9pt;color:#555\">Generata con AI (" + WebUtility.HtmlEncode(result.ActualModel) + ") - verificare prima dell'invio." +
                (warnings.Length > 0 ? "<br><b>Attenzione:</b> " + WebUtility.HtmlEncode(warnings) : "") + "</span></div>" +
                MarkdownHtml.Convert(result.Text) + "</div>";
            draft.Display(false);
        }

        // ---------------- risposte ----------------

        private static async Task ReplyCoreAsync(Outlook.Application app, string intent)
        {
            if (busy) return;
            var mails = SelectedMails(app.ActiveExplorer());
            if (mails.Count == 0)
            {
                MessageBox.Show("Seleziona prima un'email.", Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var mail = mails[0];
            var body = mail.Body ?? "";
            string warning = "";
            if (body.Length > MaxCharsTotal) { body = body.Substring(0, MaxCharsTotal) + "\n[... email troncata ...]"; warning = "L'email originale era molto lunga ed e' stata troncata."; }
            var context = "Da: " + mail.SenderName + "\nOggetto: " + mail.Subject + "\n\n" + body;

            var result = await RunAsync("Scrittura della risposta...",
                "Scrivi solo il corpo di una risposta a questa email (senza oggetto e senza firma), con questo intento: " + intent, context);
            if (result == null) return;

            var reply = mail.Reply();
            var block = "<div style=\"font-family:Aptos,Calibri,sans-serif;font-size:11pt;color:#000;line-height:1.5\">" +
                        MarkdownHtml.Convert(result.Text) + "</div><br>";
            reply.HTMLBody = MarkdownHtml.InsertAfterBody(reply.HTMLBody, block);
            reply.Display(false);
            if (warning.Length > 0) MessageBox.Show(warning, Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ReplyCustom(Outlook.Application app)
        {
            var text = Microsoft.VisualBasic.Interaction.InputBox("Cosa vuoi rispondere a questa email?", Title, "Conferma e chiedi maggiori dettagli...");
            if (!string.IsNullOrWhiteSpace(text)) Reply(app, text.Trim());
        }

        // ---------------- comune ----------------

        private static List<Outlook.MailItem> SelectedMails(Outlook.Explorer exp)
        {
            var list = new List<Outlook.MailItem>();
            if (exp == null) return list;
            Outlook.Selection sel;
            try { sel = exp.Selection; } catch { return list; }
            for (int i = 1; i <= sel.Count; i++)
            {
                var m = sel[i] as Outlook.MailItem;
                if (m != null) list.Add(m);
            }
            return list;
        }

        /// <summary>Chiamata al modello con finestra di avanzamento; null se annullata o in errore (gia' mostrato).</summary>
        private static async Task<AiResult> RunAsync(string message, string prompt, string context)
        {
            var settings = AppSettings.Load();
            if (!settings.HasAnyKey())
            {
                using (var f = new SettingsForm(settings, "Serve una TUA chiave API personale (non inclusa), legata al tuo contratto con il provider. " +
                                                          "La piu' semplice: Google Gemini -> \"Ottieni chiave\"."))
                    if (f.ShowDialog() != DialogResult.OK) return null;
                settings = AppSettings.Load();
            }
            var provider = settings.Provider;
            var model = settings.ModelFor(provider);

            busy = true;
            using (var cts = new CancellationTokenSource())
            using (var progress = new ProgressForm(message + "\n" + ModelCatalog.ProviderNames[provider] + " - " + model, cts))
            {
                progress.Show();
                try
                {
                    return await AiClient.CallAsync(provider, model, settings.GetKey(provider), SystemPrompt, prompt, context, cts.Token);
                }
                catch (OperationCanceledException) { return null; }
                catch (AiException ex)
                {
                    progress.Hide();
                    MessageBox.Show(ex.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return null;
                }
                catch (Exception ex)
                {
                    progress.Hide();
                    MessageBox.Show("Errore imprevisto: " + ex.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                finally
                {
                    busy = false;
                    progress.Close();
                }
            }
        }
    }
}
