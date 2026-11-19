using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace PersonalAIAssistant
{
    /// <summary>Impostazioni: provider, modello per provider e chiavi API (salvate cifrate).</summary>
    internal sealed class SettingsForm : Form
    {
        private readonly AppSettings settings;
        private readonly ComboBox providerBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox modelBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly Dictionary<string, TextBox> keyBoxes = new Dictionary<string, TextBox>();
        private readonly Label status = new Label { AutoSize = false, ForeColor = Color.DimGray };
        private readonly Dictionary<string, string> modelChoice = new Dictionary<string, string>();

        public SettingsForm(AppSettings settings, string intro)
        {
            this.settings = settings;
            Text = "Personal AI Assistant - Impostazioni";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false; MaximizeBox = false;
            Font = new Font("Segoe UI", 9F);
            ClientSize = new Size(520, 400);

            int y = 12;
            if (!string.IsNullOrEmpty(intro))
            {
                Controls.Add(new Label { Text = intro, Location = new Point(14, y), Size = new Size(492, 48), ForeColor = Color.FromArgb(26, 115, 232) });
                y += 52;
            }

            Controls.Add(new Label { Text = "Provider:", Location = new Point(14, y + 4), AutoSize = true });
            providerBox.Location = new Point(110, y); providerBox.Size = new Size(396, 24);
            foreach (var kv in ModelCatalog.ProviderNames) providerBox.Items.Add(kv.Value);
            Controls.Add(providerBox); y += 32;

            Controls.Add(new Label { Text = "Modello:", Location = new Point(14, y + 4), AutoSize = true });
            modelBox.Location = new Point(110, y); modelBox.Size = new Size(396, 24);
            Controls.Add(modelBox); y += 40;

            Controls.Add(new Label { Text = "Chiavi API (cifrate per il tuo utente Windows, inviate solo al provider scelto):", Location = new Point(14, y), Size = new Size(492, 18) });
            y += 22;
            foreach (var kv in ModelCatalog.ProviderNames)
            {
                var p = kv.Key;
                Controls.Add(new Label { Text = kv.Value + ":", Location = new Point(14, y + 4), AutoSize = true });
                var tb = new TextBox { UseSystemPasswordChar = true, Location = new Point(150, y), Size = new Size(250, 24), Text = settings.GetKey(p) };
                keyBoxes[p] = tb; Controls.Add(tb);
                var link = new LinkLabel { Text = "Ottieni chiave", Location = new Point(408, y + 4), AutoSize = true };
                var url = ModelCatalog.KeyUrls[p];
                link.LinkClicked += (s, e) => { try { Process.Start(url); } catch { } };
                Controls.Add(link);
                y += 30;
            }

            status.Location = new Point(14, y + 6); status.Size = new Size(492, 60); Controls.Add(status);

            var verify = new Button { Text = "Verifica chiavi", Location = new Point(14, 360), Size = new Size(110, 28) };
            var clear = new Button { Text = "Cancella chiavi", Location = new Point(130, 360), Size = new Size(110, 28) };
            var ok = new Button { Text = "Salva", Location = new Point(330, 360), Size = new Size(84, 28) };
            var cancel = new Button { Text = "Annulla", Location = new Point(422, 360), Size = new Size(84, 28), DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { verify, clear, ok, cancel });
            AcceptButton = ok; CancelButton = cancel;

            foreach (var p in ModelCatalog.ProviderNames.Keys) modelChoice[p] = settings.ModelFor(p);
            providerBox.SelectedIndex = Math.Max(0, ModelCatalog.ProviderNames.Keys.ToList().IndexOf(settings.Provider));
            FillModels();
            providerBox.SelectedIndexChanged += (s, e) => FillModels();
            modelBox.SelectedIndexChanged += (s, e) =>
            {
                var m = modelBox.SelectedItem as ModelCatalog.Model;
                if (m != null) modelChoice[m.Provider] = m.Id;
            };

            ok.Click += (s, e) => { SaveAndClose(); };
            clear.Click += (s, e) =>
            {
                foreach (var tb in keyBoxes.Values) tb.Text = "";
                status.Text = "Chiavi cancellate: premi Salva per confermare.";
            };
            verify.Click += async (s, e) =>
            {
                verify.Enabled = false; status.Text = "Verifica in corso...";
                var lines = new List<string>();
                using (var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30)))
                {
                    foreach (var kv in keyBoxes)
                    {
                        var key = kv.Value.Text.Trim();
                        var name = ModelCatalog.ProviderNames[kv.Key];
                        if (key.Length == 0) { lines.Add(name + ": nessuna chiave"); continue; }
                        try
                        {
                            var ids = await AiClient.ListModelsAsync(kv.Key, key, cts.Token);
                            var missing = ModelCatalog.Models.Where(m => m.Provider == kv.Key && !ids.Contains(m.Id)).Select(m => m.Id).ToList();
                            lines.Add(name + ": valida" + (missing.Count > 0 ? " (modelli non trovati: " + string.Join(", ", missing) + ")" : ""));
                        }
                        catch (Exception ex) { lines.Add(name + ": " + ex.Message.Split('\n')[0]); }
                    }
                }
                status.Text = string.Join("\n", lines);
                verify.Enabled = true;
            };
        }

        private string SelectedProvider { get { return ModelCatalog.ProviderNames.Keys.ElementAt(Math.Max(0, providerBox.SelectedIndex)); } }

        private void FillModels()
        {
            var p = SelectedProvider;
            modelBox.Items.Clear();
            foreach (var m in ModelCatalog.Models.Where(m => m.Provider == p)) modelBox.Items.Add(m);
            var current = ModelCatalog.Find(p, modelChoice[p]);
            modelBox.SelectedItem = current ?? modelBox.Items[0];
        }

        private void SaveAndClose()
        {
            foreach (var kv in keyBoxes) settings.SetKey(kv.Key, kv.Value.Text);
            settings.Provider = SelectedProvider;
            foreach (var kv in modelChoice) settings.Models[kv.Key] = kv.Value;
            if (settings.GetKey(settings.Provider).Length == 0)
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
