using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace UNF_MailFile_GDUpload
{
    public partial class SettingsForm : Form
    {
        private readonly GoogleDriveService googleDriveService;

        public SettingsForm(GoogleDriveService googleDriveService)
        {
            if (googleDriveService == null)
            {
                throw new ArgumentNullException(nameof(googleDriveService));
            }

            this.googleDriveService = googleDriveService;
            this.InitializeComponent();
            this.ApplyDarkTheme();
            this.LoadCurrentSettings();
        }

        private void LoadCurrentSettings()
        {
            this.txtRootFolderId.Text = Properties.Settings.Default.GoogleDriveRootFolderId ?? string.Empty;
            this.txtTokenPath.Text = this.googleDriveService.GetTokenStoreDirectoryPath();
            this.chkEnablePasswordZipWorkflow.Checked = Properties.Settings.Default.EnablePasswordZipWorkflow;
            this.chkZipUseAesEncryption.Checked = Properties.Settings.Default.ZipUseAesEncryption;
            this.chkZipValidateAfterCreate.Checked = Properties.Settings.Default.ZipValidateAfterCreate;
            this.chkZipNormalizeEntryNameToAscii.Checked = Properties.Settings.Default.ZipNormalizeEntryNameToAscii;

            this.lblClientConfigStatusValue.Text = this.googleDriveService.HasRequiredClientConfiguration()
                ? "client_id / client_secret 設定済み"
                : "client_id / client_secret 未設定";
            this.lblClientConfigStatusValue.ForeColor = this.googleDriveService.HasRequiredClientConfiguration()
                ? Color.FromArgb(140, 220, 140)
                : Color.FromArgb(255, 170, 120);

            if (this.googleDriveService.HasStoredToken())
            {
                this.lblAuthorizationStatusValue.Text = "認証済み";
                this.lblAuthorizationStatusValue.ForeColor = Color.FromArgb(140, 220, 140);
            }
            else
            {
                this.lblAuthorizationStatusValue.Text = "未認証";
                this.lblAuthorizationStatusValue.ForeColor = Color.FromArgb(255, 170, 120);
            }
        }

        private async void btnAuthenticate_Click(object sender, EventArgs e)
        {
            this.ToggleBusyState(true, "Google OAuth 2.0 フローを開始しています...");

            try
            {
                await this.googleDriveService.AuthorizeAsync(CancellationToken.None);
                this.lblAuthorizationStatusValue.Text = "認証済み";
                this.lblAuthorizationStatusValue.ForeColor = Color.FromArgb(140, 220, 140);
                MessageBox.Show(
                    "Google Drive の認証が正常に完了しました。",
                    "認証完了",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                this.lblAuthorizationStatusValue.Text = "認証失敗";
                this.lblAuthorizationStatusValue.ForeColor = Color.FromArgb(255, 170, 120);
                MessageBox.Show(
                    "Google Drive の認証に失敗しました。" + Environment.NewLine + ex.Message,
                    "認証エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                this.ToggleBusyState(false, "Google Drive の OAuth 認証をここから実行してください。");
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.GoogleDriveRootFolderId = this.txtRootFolderId.Text.Trim();
            Properties.Settings.Default.EnablePasswordZipWorkflow = this.chkEnablePasswordZipWorkflow.Checked;
            Properties.Settings.Default.ZipUseAesEncryption = this.chkZipUseAesEncryption.Checked;
            Properties.Settings.Default.ZipValidateAfterCreate = this.chkZipValidateAfterCreate.Checked;
            Properties.Settings.Default.ZipNormalizeEntryNameToAscii = this.chkZipNormalizeEntryNameToAscii.Checked;
            Properties.Settings.Default.Save();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnOpenLog_Click(object sender, EventArgs e)
        {
            if (System.IO.File.Exists(ThisAddIn.DiagnosticLogPath))
            {
                System.Diagnostics.Process.Start(ThisAddIn.DiagnosticLogPath);
            }
            else
            {
                MessageBox.Show(
                    "診断ログがまだ作成されていません。" + Environment.NewLine +
                    "Outlook でメールを送信すると記録が開始されます。" + Environment.NewLine + Environment.NewLine +
                    "保存先: " + ThisAddIn.DiagnosticLogPath,
                    "ログなし",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void ToggleBusyState(bool isBusy, string statusText)
        {
            this.UseWaitCursor = isBusy;
            this.btnAuthenticate.Enabled = !isBusy;
            this.btnSave.Enabled = !isBusy;
            this.btnCancel.Enabled = !isBusy;
            this.btnOpenLog.Enabled = !isBusy;
            this.chkEnablePasswordZipWorkflow.Enabled = !isBusy;
            this.chkZipUseAesEncryption.Enabled = !isBusy;
            this.chkZipValidateAfterCreate.Enabled = !isBusy;
            this.chkZipNormalizeEntryNameToAscii.Enabled = !isBusy;
            this.txtRootFolderId.Enabled = !isBusy;
            this.lblDescription.Text = statusText;
        }

        private void ApplyDarkTheme()
        {
            Color backgroundColor = ColorTranslator.FromHtml("#2D2D2D");
            Color surfaceColor = ColorTranslator.FromHtml("#3A3A3A");
            Color accentColor = ColorTranslator.FromHtml("#4C8DDA");
            Color textColor = ColorTranslator.FromHtml("#E0E0E0");
            Color mutedTextColor = ColorTranslator.FromHtml("#B8B8B8");
            Color borderColor = ColorTranslator.FromHtml("#4A4A4A");

            this.BackColor = backgroundColor;
            this.ForeColor = textColor;
            this.pnlHeader.BackColor = backgroundColor;
            this.pnlBody.BackColor = surfaceColor;
            this.pnlFooter.BackColor = backgroundColor;

            foreach (Control control in this.Controls)
            {
                this.ApplyThemeRecursive(control, backgroundColor, surfaceColor, accentColor, textColor, mutedTextColor, borderColor);
            }
        }

        private void ApplyThemeRecursive(Control control, Color backgroundColor, Color surfaceColor, Color accentColor, Color textColor, Color mutedTextColor, Color borderColor)
        {
            if (control is Panel panel)
            {
                panel.BackColor = panel == this.pnlHeader || panel == this.pnlFooter ? backgroundColor : surfaceColor;
            }
            else if (control is TextBox textBox)
            {
                textBox.BackColor = backgroundColor;
                textBox.ForeColor = textColor;
                textBox.BorderStyle = BorderStyle.FixedSingle;
            }
            else if (control is Button button)
            {
                button.BackColor = accentColor;
                button.ForeColor = textColor;
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderColor = borderColor;
                button.FlatAppearance.MouseOverBackColor = Color.FromArgb(82, 122, 188);
                button.FlatAppearance.MouseDownBackColor = Color.FromArgb(68, 105, 163);
            }
            else if (control is Label label)
            {
                label.ForeColor = label == this.lblDescription ? mutedTextColor : textColor;
                label.BackColor = Color.Transparent;
            }
            else if (control is CheckBox checkBox)
            {
                checkBox.ForeColor = textColor;
                checkBox.BackColor = Color.Transparent;
            }
            else if (control is GroupBox groupBox)
            {
                groupBox.ForeColor = textColor;
                groupBox.BackColor = surfaceColor;
            }

            foreach (Control childControl in control.Controls)
            {
                this.ApplyThemeRecursive(childControl, backgroundColor, surfaceColor, accentColor, textColor, mutedTextColor, borderColor);
            }
        }
    }
}
