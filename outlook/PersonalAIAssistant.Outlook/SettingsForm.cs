using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace PersonalAIAssistant
{
    /// <summary>
    /// Impostazioni: provider; per il provider scelto modello (dal catalogo o scritto a mano), indirizzo API
    /// (solo provider compatibili OpenAI, modificabile) e chiave (salvata cifrata).
    /// </summary>
    internal sealed class SettingsForm : Form
    {
        private readonly AppSettings settings;
        private readonly ComboBox providerBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox modelBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown };
        private readonly TextBox urlBox = new TextBox();
        private readonly LinkLabel urlReset = new LinkLabel { Text = "Predefinito", AutoSize = true };
        private readonly TextBox keyBox = new TextBox { UseSystemPasswordChar = true };
        private readonly LinkLabel keyLink = new LinkLabel { Text = "Ottieni chiave", AutoSize = true };
        private readonly Label keysSummary = new Label { AutoSize = false, ForeColor = Color.DimGray };
        private readonly Label status = new Label { AutoSize = false, ForeColor = Color.DimGray };

        // Bozze per provider: si salvano tutte insieme con "Salva".
        private readonly Dictionary<string, string> keyDraft = new Dictionary<string, string>();
        private readonly Dictionary<string, string> modelDraft = new Dictionary<string, string>();
        private readonly Dictionary<string, string> urlDraft = new Dictionary<string, string>();
        private string shownProvider;

        public SettingsForm(AppSettings settings, string intro)
        {
            this.settings = settings;
            Text = "Personal AI Assistant - Impostazioni";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false; MaximizeBox = false;
            Font = new Font("Segoe UI", 9F);
            ClientSize = new Size(540, 470);

            int y = 12;
            if (!string.IsNullOrEmpty(intro))
            {
                Controls.Add(new Label { Text = intro, Location = new Point(14, y), Size = new Size(512, 48), ForeColor = Color.FromArgb(26, 115, 232) });
                y += 52;
            }

            Controls.Add(new Label { Text = "Provider:", Location = new Point(14, y + 4), AutoSize = true });
            providerBox.Location = new Point(120, y); providerBox.Size = new Size(406, 24);
            foreach (var kv in ModelCatalog.ProviderNames) providerBox.Items.Add(kv.Value);
            Controls.Add(providerBox); y += 34;

            Controls.Add(new Label { Text = "Modello:", Location = new Point(14, y + 4), AutoSize = true });
            modelBox.Location = new Point(120, y); modelBox.Size = new Size(406, 24);
            Controls.Add(modelBox); y += 26;
            Controls.Add(new Label { Text = "Scegli dall'elenco o scrivi l'ID esatto di un altro modello del provider.",
                                     Location = new Point(120, y), Size = new Size(406, 16), ForeColor = Color.DimGray }); y += 24;

            Controls.Add(new Label { Text = "Indirizzo API:", Location = new Point(14, y + 4), AutoSize = true });
            urlBox.Location = new Point(120, y); urlBox.Size = new Size(330, 24);
            urlReset.Location = new Point(458, y + 4);
            Controls.Add(urlBox); Controls.Add(urlReset); y += 26;
            Controls.Add(new Label { Text = "Provider su Internet: cambialo solo se il provider lo ha cambiato. Locale: Ollama :11434/v1, llama-server :8080/v1, LM Studio :1234/v1, oppure l'IP di un PC della rete.",
                                     Location = new Point(120, y), Size = new Size(406, 16), ForeColor = Color.DimGray }); y += 26;

            Controls.Add(new Label { Text = "Chiave API:", Location = new Point(14, y + 4), AutoSize = true });
            keyBox.Location = new Point(120, y); keyBox.Size = new Size(320, 24);
            keyLink.Location = new Point(448, y + 4);
            Controls.Add(keyBox); Controls.Add(keyLink); y += 28;
            Controls.Add(new Label { Text = "Cifrate per il tuo utente Windows; inviate solo all'indirizzo API di quel provider (se lo cambi, ti chiedo conferma).",
                                     Location = new Point(120, y), Size = new Size(406, 16), ForeColor = Color.DimGray }); y += 24;

            keysSummary.Location = new Point(14, y); keysSummary.Size = new Size(512, 34); Controls.Add(keysSummary); y += 36;
            status.Location = new Point(14, y); status.Size = new Size(512, 430 - y); Controls.Add(status);

            var verify = new Button { Text = "Verifica chiavi", Location = new Point(14, 430), Size = new Size(110, 28) };
            var clear = new Button { Text = "Cancella chiavi", Location = new Point(130, 430), Size = new Size(110, 28) };
            var ok = new Button { Text = "Salva", Location = new Point(350, 430), Size = new Size(84, 28) };
            var cancel = new Button { Text = "Annulla", Location = new Point(442, 430), Size = new Size(84, 28), DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { verify, clear, ok, cancel });
            AcceptButton = ok; CancelButton = cancel;

            foreach (var p in ModelCatalog.ProviderNames.Keys)
            {
                keyDraft[p] = settings.GetKey(p);
                modelDraft[p] = settings.ModelFor(p);
                if (ModelCatalog.IsCompatible(p)) urlDraft[p] = settings.BaseUrlFor(p);
            }
            providerBox.SelectedIndex = Math.Max(0, ModelCatalog.ProviderNames.Keys.ToList().IndexOf(settings.Provider));
            ShowProvider();
            providerBox.SelectedIndexChanged += (s, e) => { StoreShown(); ShowProvider(); };
            keyBox.TextChanged += (s, e) => { if (shownProvider != null) { keyDraft[shownProvider] = keyBox.Text; UpdateSummary(); } };
            urlReset.LinkClicked += (s, e) => { urlBox.Text = ModelCatalog.CompatBaseUrls[shownProvider]; };
            keyLink.LinkClicked += (s, e) => { try { Process.Start(ModelCatalog.KeyUrls[shownProvider]); } catch { } };

            ok.Click += (s, e) => { SaveAndClose(); };
            clear.Click += (s, e) =>
            {
                foreach (var p in keyDraft.Keys.ToList()) keyDraft[p] = "";
                keyBox.Text = "";
                UpdateSummary();
                status.Text = "Chiavi cancellate: premi Salva per confermare.";
            };
            verify.Click += async (s, e) =>
            {
                StoreShown();
                verify.Enabled = false; status.Text = "Verifica in corso...";
                var lines = new List<string>();
                using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60)))
                {
                    foreach (var p in ModelCatalog.ProviderNames.Keys)
                    {
                        var key = (keyDraft[p] ?? "").Trim();
                        if (key.Length == 0 && !(p == ModelCatalog.Local && SelectedProvider == p)) continue;
                        var name = ModelCatalog.ProviderNames[p];
                        // Si verifica solo cio' che e' gia' salvato (e quindi confermato): nessun invio prima delle conferme.
                        bool compat = ModelCatalog.IsCompatible(p);
                        if (settings.RecoveredFrom != null || key != settings.GetKey(p) ||
                            (compat && ModelCatalog.NormalizeBaseUrl(urlDraft[p], p) != settings.BaseUrlFor(p)) ||
                            !settings.IsDestinationApproved(p))
                        {
                            lines.Add(name + ": modifiche non ancora salvate - premi Salva (con le eventuali conferme), poi Verifica.");
                            continue;
                        }
                        try
                        {
                            var ids = await AiClient.ListModelsAsync(p, key, cts.Token, ModelCatalog.IsCompatible(p) ? urlDraft[p] : null);
                            var model = modelDraft[p];
                            lines.Add(name + ": valida" + (ids.Contains(model) ? "" : " (il modello " + model + " non risulta nell'elenco del provider)"));
                        }
                        catch (Exception ex) { lines.Add(name + ": " + ex.Message.Split('\n')[0]); }
                    }
                }
                status.Text = lines.Count > 0 ? string.Join("\n", lines) : "Nessuna chiave da verificare.";
                verify.Enabled = true;
            };
        }

        private string SelectedProvider { get { return ModelCatalog.ProviderNames.Keys.ElementAt(Math.Max(0, providerBox.SelectedIndex)); } }

        /// <summary>Copia nei draft quello che e' scritto nei campi del provider mostrato.</summary>
        private void StoreShown()
        {
            if (shownProvider == null) return;
            var m = modelBox.SelectedItem as ModelCatalog.Model;
            modelDraft[shownProvider] = m != null && modelBox.Text == m.Label ? m.Id : modelBox.Text.Trim();
            keyDraft[shownProvider] = keyBox.Text;
            if (ModelCatalog.IsCompatible(shownProvider)) urlDraft[shownProvider] = urlBox.Text.Trim();
        }

        private void ShowProvider()
        {
            var p = SelectedProvider;
            shownProvider = null;   // evita che gli eventi dei campi scrivano nel provider sbagliato
            modelBox.Items.Clear();
            foreach (var m in ModelCatalog.Models.Where(m => m.Provider == p)) modelBox.Items.Add(m);
            var current = ModelCatalog.Find(p, modelDraft[p]);
            if (current != null) modelBox.SelectedItem = current; else modelBox.Text = modelDraft[p];

            bool compat = ModelCatalog.IsCompatible(p);
            urlBox.Enabled = compat; urlReset.Visible = compat;
            urlBox.Text = compat ? urlDraft[p] : "(fisso per questo provider)";
            bool local = p == ModelCatalog.Local;
            keyBox.Text = keyDraft[p];
            keyLink.Text = local ? "Scarica Ollama" : "Ottieni chiave";
            shownProvider = p;
            UpdateSummary();
        }

        private void UpdateSummary()
        {
            var with = ModelCatalog.ProviderNames.Where(kv => kv.Key != ModelCatalog.Local && (keyDraft[kv.Key] ?? "").Trim().Length > 0).Select(kv => kv.Value).ToList();
            if (shownProvider == ModelCatalog.Local)
            {
                keysSummary.Text = "Chiave facoltativa (solo se il server la richiede). Il server deve essere avviato e il modello scaricato. " +
                                   "Le email restano nella tua rete; con http:// verso un altro PC viaggiano in chiaro sulla rete interna.";
                return;
            }
            keysSummary.Text = with.Count == 0 ? "Nessuna chiave inserita: senza una chiave Personal AI non puo' elaborare nulla."
                                               : "Chiavi inserite: " + string.Join(", ", with) + ".";
        }

        private void SaveAndClose()
        {
            StoreShown();

            // Controlli prima di toccare le impostazioni.
            foreach (var p in ModelCatalog.ProviderNames.Keys)
            {
                var m = modelDraft[p] ?? "";
                if (m.Length == 0 || m.Length > 200 || m.Any(char.IsWhiteSpace))
                {
                    MessageBox.Show(this, "Modello non valido per " + ModelCatalog.ProviderNames[p] + ": scrivi l'ID senza spazi.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (ModelCatalog.IsCompatible(p) && ModelCatalog.NormalizeBaseUrl(urlDraft[p], p) == null)
                {
                    MessageBox.Show(this, "Indirizzo API non valido per " + ModelCatalog.ProviderNames[p] +
                        (p == ModelCatalog.Local ? ": per il modello locale deve puntare a questo PC o alla rete interna (es. http://localhost:11434/v1 o http://192.168.1.20:11434/v1)."
                                                 : ": deve iniziare con https://.") + " Usa \"Predefinito\" per ripristinarlo.",
                        Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // Indirizzo non ufficiale con una chiave: conferma esplicita del server destinatario, legata a server + chiave.
            var approvals = new Dictionary<string, string>(settings.KeyApprovals);
            foreach (var p in ModelCatalog.CompatBaseUrls.Keys)
            {
                var key = (keyDraft[p] ?? "").Trim();
                var url = ModelCatalog.NormalizeBaseUrl(urlDraft[p], p);
                if (AppSettings.DestinationApproved(p, url, key, approvals)) continue;
                var origin = AppSettings.OriginOf(url);
                var msg = "La chiave di " + ModelCatalog.ProviderNames[p] + " verra' inviata a:\n\n    " + origin + "\n\n" +
                          "invece che all'indirizzo ufficiale (" + AppSettings.OriginOf(ModelCatalog.CompatBaseUrls[p]) + "). " +
                          "Chi gestisce quel server ricevera' la chiave e i testi delle email che elabori.\n\n" +
                          "Confermi solo se ti fidi di questo server. Continuare?";
                if (MessageBox.Show(this, msg, "Indirizzo API non ufficiale", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                                    MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                {
                    status.Text = "Impostazioni non salvate: usa \"Predefinito\" per tornare all'indirizzo ufficiale.";
                    return;
                }
                approvals[p] = AppSettings.ApprovalToken(origin, key);
            }

            // Chiavi nuove o cambiate di provider con sede in Cina: conferma esplicita.
            var china = ModelCatalog.ChinaProviders.Keys
                .Where(p => (keyDraft[p] ?? "").Trim().Length > 0 && (keyDraft[p] ?? "").Trim() != settings.GetKey(p)).ToList();
            if (china.Count > 0)
            {
                var msg = "Con " + string.Join(", ", china.Select(p => ModelCatalog.ProviderNames[p])) +
                          " le email che elabori verranno inviate ai server di aziende con sede in Cina e trattate secondo le loro " +
                          "condizioni e le leggi applicabili, anche cinesi.\n\n" +
                          string.Join("\n", china.Select(p => "- " + ModelCatalog.ChinaProviders[p])) +
                          "\n\nNon usarli per dati personali, riservati o aziendali se non sei autorizzato a farlo.\n\nSalvare comunque?";
                if (MessageBox.Show(this, msg, "Dati inviati a server di aziende cinesi", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                                    MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                {
                    status.Text = "Impostazioni non salvate.";
                    return;
                }
            }

            foreach (var p in ModelCatalog.ProviderNames.Keys)
            {
                settings.SetKey(p, keyDraft[p]);
                settings.Models[p] = modelDraft[p];
                if (ModelCatalog.IsCompatible(p))
                {
                    var u = ModelCatalog.NormalizeBaseUrl(urlDraft[p], p);
                    if (u == ModelCatalog.CompatBaseUrls[p]) settings.BaseUrls.Remove(p); else settings.BaseUrls[p] = u;
                }
            }
            // Restano solo i consensi ancora validi per server + chiave salvati.
            settings.KeyApprovals.Clear();
            foreach (var p in ModelCatalog.CompatBaseUrls.Keys)
            {
                string t;
                var origin = AppSettings.OriginOf(settings.BaseUrlFor(p));
                if (approvals.TryGetValue(p, out t) && t == AppSettings.ApprovalToken(origin, settings.GetKey(p)) &&
                    origin != AppSettings.OriginOf(ModelCatalog.CompatBaseUrls[p]))
                    settings.KeyApprovals[p] = t;
            }
            settings.Provider = SelectedProvider;
            if (settings.Provider != ModelCatalog.Local && settings.GetKey(settings.Provider).Length == 0)
            {
                var withKey = ModelCatalog.ProviderNames.Keys.FirstOrDefault(p => settings.GetKey(p).Length > 0);
                if (withKey != null &&
                    MessageBox.Show(this, "Non hai una chiave per " + ModelCatalog.ProviderNames[settings.Provider] +
                        ". Vuoi usare " + ModelCatalog.ProviderNames[withKey] + ", per cui hai una chiave?",
                        Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    settings.Provider = withKey;
            }
            try { settings.Save(); }
            catch (Exception ex) { MessageBox.Show(this, "Impossibile salvare le impostazioni: " + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
