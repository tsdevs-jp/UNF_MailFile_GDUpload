using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace UNF_MailFile_GDUpload
{
    internal sealed class UploadProgressForm : Form
    {
        private readonly Label lblStatus;
        private readonly ProgressBar progressBar;

        internal UploadProgressForm()
        {
            this.Text = "Google Drive アップロード中";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(440, 104);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;
            this.ShowInTaskbar = false;

            this.progressBar = new ProgressBar
            {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Location = new Point(16, 14),
                Size = new Size(408, 22),
                TabStop = false
            };

            this.lblStatus = new Label
            {
                Text = "処理を開始しています...",
                Location = new Point(16, 52),
                Size = new Size(408, 38),
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                UseMnemonic = false
            };

            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
        }

        internal void UpdateStatus(string message)
        {
            if (this.IsDisposed) return;
            if (this.InvokeRequired)
            {
                try { this.BeginInvoke(new Action<string>(this.UpdateStatus), message); } catch { }
                return;
            }
            this.lblStatus.Text = message ?? string.Empty;
        }
    }
}
