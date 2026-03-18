using System.ComponentModel;
using System.Windows.Forms;

namespace UNF_MailFile_GDUpload
{
    partial class SettingsForm
    {
        private IContainer components = null;
        private Panel pnlHeader;
        private Panel pnlBody;
        private Panel pnlFooter;
        private Label lblTitle;
        private Label lblDescription;
        private GroupBox grpGoogleAuth;
        private Label lblAuthorizationStatus;
        private Label lblAuthorizationStatusValue;
        private Label lblTokenPath;
        private TextBox txtTokenPath;
        private Button btnAuthenticate;
        private GroupBox grpFolder;
        private Label lblRootFolderId;
        private TextBox txtRootFolderId;
        private Label lblRootFolderHint;
        private GroupBox grpClientConfig;
        private Label lblClientConfigStatus;
        private Label lblClientConfigStatusValue;
        private Label lblClientConfigHint;
        private Button btnSave;
        private Button btnCancel;
        private Button btnOpenLog;
        private CheckBox chkEnablePasswordZipWorkflow;
        private CheckBox chkZipUseAesEncryption;
        private CheckBox chkZipValidateAfterCreate;
        private CheckBox chkZipNormalizeEntryNameToAscii;

        protected override void Dispose(bool disposing)
        {
            if (disposing && this.components != null)
            {
                this.components.Dispose();
            }

            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.pnlHeader = new System.Windows.Forms.Panel();
            this.lblDescription = new System.Windows.Forms.Label();
            this.lblTitle = new System.Windows.Forms.Label();
            this.pnlBody = new System.Windows.Forms.Panel();
            this.grpClientConfig = new System.Windows.Forms.GroupBox();
            this.lblClientConfigHint = new System.Windows.Forms.Label();
            this.lblClientConfigStatusValue = new System.Windows.Forms.Label();
            this.lblClientConfigStatus = new System.Windows.Forms.Label();
            this.grpFolder = new System.Windows.Forms.GroupBox();
            this.lblRootFolderHint = new System.Windows.Forms.Label();
            this.txtRootFolderId = new System.Windows.Forms.TextBox();
            this.lblRootFolderId = new System.Windows.Forms.Label();
            this.grpGoogleAuth = new System.Windows.Forms.GroupBox();
            this.btnAuthenticate = new System.Windows.Forms.Button();
            this.txtTokenPath = new System.Windows.Forms.TextBox();
            this.lblTokenPath = new System.Windows.Forms.Label();
            this.lblAuthorizationStatusValue = new System.Windows.Forms.Label();
            this.lblAuthorizationStatus = new System.Windows.Forms.Label();
            this.pnlFooter = new System.Windows.Forms.Panel();
            this.chkZipNormalizeEntryNameToAscii = new System.Windows.Forms.CheckBox();
            this.chkZipValidateAfterCreate = new System.Windows.Forms.CheckBox();
            this.chkZipUseAesEncryption = new System.Windows.Forms.CheckBox();
            this.chkEnablePasswordZipWorkflow = new System.Windows.Forms.CheckBox();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnOpenLog = new System.Windows.Forms.Button();
            this.pnlHeader.SuspendLayout();
            this.pnlBody.SuspendLayout();
            this.grpClientConfig.SuspendLayout();
            this.grpFolder.SuspendLayout();
            this.grpGoogleAuth.SuspendLayout();
            this.pnlFooter.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlHeader
            // 
            this.pnlHeader.Controls.Add(this.lblDescription);
            this.pnlHeader.Controls.Add(this.lblTitle);
            this.pnlHeader.Dock = System.Windows.Forms.DockStyle.Top;
            this.pnlHeader.Location = new System.Drawing.Point(0, 0);
            this.pnlHeader.Name = "pnlHeader";
            this.pnlHeader.Padding = new System.Windows.Forms.Padding(24, 20, 24, 16);
            this.pnlHeader.Size = new System.Drawing.Size(760, 96);
            this.pnlHeader.TabIndex = 0;
            // 
            // lblDescription
            // 
            this.lblDescription.AutoSize = true;
            this.lblDescription.Location = new System.Drawing.Point(27, 53);
            this.lblDescription.Name = "lblDescription";
            this.lblDescription.Size = new System.Drawing.Size(330, 19);
            this.lblDescription.TabIndex = 1;
            this.lblDescription.Text = "Google Drive の OAuth 認証をここから実行してください。";
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 16F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.lblTitle.Location = new System.Drawing.Point(24, 16);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(295, 30);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "Google Drive アップロード設定";
            // 
            // pnlBody
            // 
            this.pnlBody.Controls.Add(this.grpClientConfig);
            this.pnlBody.Controls.Add(this.grpFolder);
            this.pnlBody.Controls.Add(this.grpGoogleAuth);
            this.pnlBody.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pnlBody.Location = new System.Drawing.Point(0, 96);
            this.pnlBody.Name = "pnlBody";
            this.pnlBody.Padding = new System.Windows.Forms.Padding(24, 8, 24, 8);
            this.pnlBody.Size = new System.Drawing.Size(760, 430);
            this.pnlBody.TabIndex = 1;
            // 
            // grpClientConfig
            // 
            this.grpClientConfig.Controls.Add(this.lblClientConfigHint);
            this.grpClientConfig.Controls.Add(this.lblClientConfigStatusValue);
            this.grpClientConfig.Controls.Add(this.lblClientConfigStatus);
            this.grpClientConfig.Location = new System.Drawing.Point(24, 280);
            this.grpClientConfig.Name = "grpClientConfig";
            this.grpClientConfig.Padding = new System.Windows.Forms.Padding(16, 14, 16, 14);
            this.grpClientConfig.Size = new System.Drawing.Size(712, 128);
            this.grpClientConfig.TabIndex = 2;
            this.grpClientConfig.TabStop = false;
            this.grpClientConfig.Text = "クライアント設定";
            // 
            // lblClientConfigHint
            // 
            this.lblClientConfigHint.AutoSize = true;
            this.lblClientConfigHint.Location = new System.Drawing.Point(21, 75);
            this.lblClientConfigHint.Name = "lblClientConfigHint";
            this.lblClientConfigHint.Size = new System.Drawing.Size(580, 38);
            this.lblClientConfigHint.TabIndex = 2;
            this.lblClientConfigHint.Text = "App.config の appSettings に GoogleDriveClientId / GoogleDriveClientSecret を設定してくださ" +
    "い。\r\n設定はソースコード外で管理でき、環境ごとに置き換えが可能です。";
            // 
            // lblClientConfigStatusValue
            // 
            this.lblClientConfigStatusValue.AutoSize = true;
            this.lblClientConfigStatusValue.Location = new System.Drawing.Point(154, 38);
            this.lblClientConfigStatusValue.Name = "lblClientConfigStatusValue";
            this.lblClientConfigStatusValue.Size = new System.Drawing.Size(37, 19);
            this.lblClientConfigStatusValue.TabIndex = 1;
            this.lblClientConfigStatusValue.Text = "状態";
            // 
            // lblClientConfigStatus
            // 
            this.lblClientConfigStatus.AutoSize = true;
            this.lblClientConfigStatus.Location = new System.Drawing.Point(21, 38);
            this.lblClientConfigStatus.Name = "lblClientConfigStatus";
            this.lblClientConfigStatus.Size = new System.Drawing.Size(126, 19);
            this.lblClientConfigStatus.TabIndex = 0;
            this.lblClientConfigStatus.Text = "クライアント設定状態";
            // 
            // grpFolder
            // 
            this.grpFolder.Controls.Add(this.lblRootFolderHint);
            this.grpFolder.Controls.Add(this.txtRootFolderId);
            this.grpFolder.Controls.Add(this.lblRootFolderId);
            this.grpFolder.Location = new System.Drawing.Point(24, 149);
            this.grpFolder.Name = "grpFolder";
            this.grpFolder.Padding = new System.Windows.Forms.Padding(16, 14, 16, 14);
            this.grpFolder.Size = new System.Drawing.Size(712, 115);
            this.grpFolder.TabIndex = 1;
            this.grpFolder.TabStop = false;
            this.grpFolder.Text = "Drive ルートフォルダ";
            // 
            // lblRootFolderHint
            // 
            this.lblRootFolderHint.AutoSize = true;
            this.lblRootFolderHint.Location = new System.Drawing.Point(21, 77);
            this.lblRootFolderHint.Name = "lblRootFolderHint";
            this.lblRootFolderHint.Size = new System.Drawing.Size(534, 19);
            this.lblRootFolderHint.TabIndex = 2;
            this.lblRootFolderHint.Text = "送信毎に作成される [yyyyMMdd_HHmm]_[件名] フォルダの親フォルダ ID を指定してください。";
            // 
            // txtRootFolderId
            // 
            this.txtRootFolderId.Location = new System.Drawing.Point(158, 34);
            this.txtRootFolderId.Name = "txtRootFolderId";
            this.txtRootFolderId.Size = new System.Drawing.Size(530, 25);
            this.txtRootFolderId.TabIndex = 1;
            // 
            // lblRootFolderId
            // 
            this.lblRootFolderId.AutoSize = true;
            this.lblRootFolderId.Location = new System.Drawing.Point(21, 37);
            this.lblRootFolderId.Name = "lblRootFolderId";
            this.lblRootFolderId.Size = new System.Drawing.Size(98, 19);
            this.lblRootFolderId.TabIndex = 0;
            this.lblRootFolderId.Text = "ルートフォルダ ID";
            // 
            // grpGoogleAuth
            // 
            this.grpGoogleAuth.Controls.Add(this.btnAuthenticate);
            this.grpGoogleAuth.Controls.Add(this.txtTokenPath);
            this.grpGoogleAuth.Controls.Add(this.lblTokenPath);
            this.grpGoogleAuth.Controls.Add(this.lblAuthorizationStatusValue);
            this.grpGoogleAuth.Controls.Add(this.lblAuthorizationStatus);
            this.grpGoogleAuth.Location = new System.Drawing.Point(24, 16);
            this.grpGoogleAuth.Name = "grpGoogleAuth";
            this.grpGoogleAuth.Padding = new System.Windows.Forms.Padding(16, 14, 16, 14);
            this.grpGoogleAuth.Size = new System.Drawing.Size(712, 117);
            this.grpGoogleAuth.TabIndex = 0;
            this.grpGoogleAuth.TabStop = false;
            this.grpGoogleAuth.Text = "OAuth 2.0 認証";
            // 
            // btnAuthenticate
            // 
            this.btnAuthenticate.Location = new System.Drawing.Point(520, 30);
            this.btnAuthenticate.Name = "btnAuthenticate";
            this.btnAuthenticate.Size = new System.Drawing.Size(168, 34);
            this.btnAuthenticate.TabIndex = 4;
            this.btnAuthenticate.Text = "Google 認証を実行";
            this.btnAuthenticate.UseVisualStyleBackColor = true;
            this.btnAuthenticate.Click += new System.EventHandler(this.btnAuthenticate_Click);
            // 
            // txtTokenPath
            // 
            this.txtTokenPath.Location = new System.Drawing.Point(158, 75);
            this.txtTokenPath.Name = "txtTokenPath";
            this.txtTokenPath.ReadOnly = true;
            this.txtTokenPath.Size = new System.Drawing.Size(530, 25);
            this.txtTokenPath.TabIndex = 3;
            // 
            // lblTokenPath
            // 
            this.lblTokenPath.AutoSize = true;
            this.lblTokenPath.Location = new System.Drawing.Point(21, 78);
            this.lblTokenPath.Name = "lblTokenPath";
            this.lblTokenPath.Size = new System.Drawing.Size(92, 19);
            this.lblTokenPath.TabIndex = 2;
            this.lblTokenPath.Text = "トークン保存先";
            // 
            // lblAuthorizationStatusValue
            // 
            this.lblAuthorizationStatusValue.AutoSize = true;
            this.lblAuthorizationStatusValue.Location = new System.Drawing.Point(158, 37);
            this.lblAuthorizationStatusValue.Name = "lblAuthorizationStatusValue";
            this.lblAuthorizationStatusValue.Size = new System.Drawing.Size(51, 19);
            this.lblAuthorizationStatusValue.TabIndex = 1;
            this.lblAuthorizationStatusValue.Text = "未認証";
            // 
            // lblAuthorizationStatus
            // 
            this.lblAuthorizationStatus.AutoSize = true;
            this.lblAuthorizationStatus.Location = new System.Drawing.Point(21, 37);
            this.lblAuthorizationStatus.Name = "lblAuthorizationStatus";
            this.lblAuthorizationStatus.Size = new System.Drawing.Size(65, 19);
            this.lblAuthorizationStatus.TabIndex = 0;
            this.lblAuthorizationStatus.Text = "認証状態";
            // 
            // pnlFooter
            // 
            this.pnlFooter.Controls.Add(this.chkZipNormalizeEntryNameToAscii);
            this.pnlFooter.Controls.Add(this.chkZipValidateAfterCreate);
            this.pnlFooter.Controls.Add(this.chkZipUseAesEncryption);
            this.pnlFooter.Controls.Add(this.chkEnablePasswordZipWorkflow);
            this.pnlFooter.Controls.Add(this.btnCancel);
            this.pnlFooter.Controls.Add(this.btnSave);
            this.pnlFooter.Controls.Add(this.btnOpenLog);
            this.pnlFooter.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pnlFooter.Location = new System.Drawing.Point(0, 526);
            this.pnlFooter.Name = "pnlFooter";
            this.pnlFooter.Padding = new System.Windows.Forms.Padding(24, 12, 24, 12);
            this.pnlFooter.Size = new System.Drawing.Size(760, 118);
            this.pnlFooter.TabIndex = 2;
            // 
            // chkZipNormalizeEntryNameToAscii
            // 
            this.chkZipNormalizeEntryNameToAscii.AutoSize = true;
            this.chkZipNormalizeEntryNameToAscii.Location = new System.Drawing.Point(260, 40);
            this.chkZipNormalizeEntryNameToAscii.Name = "chkZipNormalizeEntryNameToAscii";
            this.chkZipNormalizeEntryNameToAscii.Size = new System.Drawing.Size(322, 23);
            this.chkZipNormalizeEntryNameToAscii.TabIndex = 6;
            this.chkZipNormalizeEntryNameToAscii.Text = "Zip内ファイル名をASCII互換名に正規化（必要時）";
            this.chkZipNormalizeEntryNameToAscii.UseVisualStyleBackColor = true;
            // 
            // chkZipValidateAfterCreate
            // 
            this.chkZipValidateAfterCreate.AutoSize = true;
            this.chkZipValidateAfterCreate.Location = new System.Drawing.Point(24, 40);
            this.chkZipValidateAfterCreate.Name = "chkZipValidateAfterCreate";
            this.chkZipValidateAfterCreate.Size = new System.Drawing.Size(209, 23);
            this.chkZipValidateAfterCreate.TabIndex = 5;
            this.chkZipValidateAfterCreate.Text = "作成後にZip整合性を自己検証";
            this.chkZipValidateAfterCreate.UseVisualStyleBackColor = true;
            // 
            // chkZipUseAesEncryption
            // 
            this.chkZipUseAesEncryption.AutoSize = true;
            this.chkZipUseAesEncryption.Location = new System.Drawing.Point(260, 12);
            this.chkZipUseAesEncryption.Name = "chkZipUseAesEncryption";
            this.chkZipUseAesEncryption.Size = new System.Drawing.Size(265, 23);
            this.chkZipUseAesEncryption.TabIndex = 4;
            this.chkZipUseAesEncryption.Text = "Zip暗号を AES-256 にする（互換性低）";
            this.chkZipUseAesEncryption.UseVisualStyleBackColor = true;
            // 
            // chkEnablePasswordZipWorkflow
            // 
            this.chkEnablePasswordZipWorkflow.AutoSize = true;
            this.chkEnablePasswordZipWorkflow.Location = new System.Drawing.Point(24, 12);
            this.chkEnablePasswordZipWorkflow.Name = "chkEnablePasswordZipWorkflow";
            this.chkEnablePasswordZipWorkflow.Size = new System.Drawing.Size(180, 23);
            this.chkEnablePasswordZipWorkflow.TabIndex = 3;
            this.chkEnablePasswordZipWorkflow.Text = "PW付きZip化を有効にする";
            this.chkEnablePasswordZipWorkflow.UseVisualStyleBackColor = true;
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(567, 68);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(169, 38);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(392, 68);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(169, 38);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnOpenLog
            // 
            this.btnOpenLog.Location = new System.Drawing.Point(24, 68);
            this.btnOpenLog.Name = "btnOpenLog";
            this.btnOpenLog.Size = new System.Drawing.Size(169, 38);
            this.btnOpenLog.TabIndex = 2;
            this.btnOpenLog.Text = "診断ログを開く";
            this.btnOpenLog.UseVisualStyleBackColor = true;
            this.btnOpenLog.Click += new System.EventHandler(this.btnOpenLog_Click);
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.btnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(760, 644);
            this.Controls.Add(this.pnlBody);
            this.Controls.Add(this.pnlFooter);
            this.Controls.Add(this.pnlHeader);
            this.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.Text = "Google Drive アップロード設定";
            this.pnlHeader.ResumeLayout(false);
            this.pnlHeader.PerformLayout();
            this.pnlBody.ResumeLayout(false);
            this.grpClientConfig.ResumeLayout(false);
            this.grpClientConfig.PerformLayout();
            this.grpFolder.ResumeLayout(false);
            this.grpFolder.PerformLayout();
            this.grpGoogleAuth.ResumeLayout(false);
            this.grpGoogleAuth.PerformLayout();
            this.pnlFooter.ResumeLayout(false);
            this.pnlFooter.PerformLayout();
            this.ResumeLayout(false);

        }
    }
}
