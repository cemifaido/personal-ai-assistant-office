using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace PersonalAIAssistant
{
    /// <summary>Finestrina non modale "Elaborazione in corso" con pulsante Annulla.</summary>
    internal sealed class ProgressForm : Form
    {
        private readonly CancellationTokenSource cts;

        public ProgressForm(string message, CancellationTokenSource cts)
        {
            this.cts = cts;
            Text = "Personal AI Assistant";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false; MaximizeBox = false; ShowInTaskbar = false; TopMost = true;
            ClientSize = new Size(380, 110);
            Font = new Font("Segoe UI", 9F);

            var label = new Label { Text = message, AutoSize = false, Location = new Point(16, 14), Size = new Size(348, 40) };
            var bar = new ProgressBar { Style = ProgressBarStyle.Marquee, Location = new Point(16, 56), Size = new Size(260, 18) };
            var cancel = new Button { Text = "Annulla", Location = new Point(288, 52), Size = new Size(76, 26) };
            cancel.Click += (s, e) => { cancel.Enabled = false; label.Text = "Annullamento..."; this.cts.Cancel(); };
            Controls.AddRange(new Control[] { label, bar, cancel });
            FormClosing += (s, e) => { if (e.CloseReason == CloseReason.UserClosing) this.cts.Cancel(); };
        }
    }
}
