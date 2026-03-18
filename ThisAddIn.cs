using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Core;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace UNF_MailFile_GDUpload
{
    public partial class ThisAddIn
    {
        private enum SendProcessingMode
        {
            GoogleDrive,
            EncryptedZipAttachment
        }

        private const string SkipProcessingUserPropertyName = "UNFGDUploadSkip";
        private const string ProcessingUserPropertyName = "UNFGDUploadProcessing";
        private const string OutlookMaximumAttachmentSizeValueName = "MaximumAttachmentSize";
        private const string ZipPasswordPrefix = "UNF-";
        private const int ZipPasswordDigits = 8;
        private const string AttachmentHiddenPropertyTag = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";
        private const string AttachmentContentIdPropertyTag = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";
        private const string AttachmentFlagsPropertyTag = "http://schemas.microsoft.com/mapi/proptag/0x37140003";
        private const int AttachmentFlagMhtmlRef = 0x00000004;

        public static readonly string DiagnosticLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "UNF_MailFile_GDUpload",
            "diagnostic.log");

        private readonly object processingLock = new object();
        private readonly Dictionary<string, MailSendOperationContext> processingOperations = new Dictionary<string, MailSendOperationContext>(StringComparer.OrdinalIgnoreCase);

        private Control marshalingControl;
        private GoogleDriveService googleDriveService;
        private CommandBarButton settingsCommandBarButton;

        private string outlookPreferencesRegistryPath;
        private object originalMaximumAttachmentSizeValue;
        private RegistryValueKind? originalMaximumAttachmentSizeValueKind;
        private bool hadMaximumAttachmentSizeValue;
        private bool attachmentLimitOverrideApplied;

        private bool IsPasswordZipWorkflowEnabled
        {
            get
            {
                return Properties.Settings.Default.EnablePasswordZipWorkflow;
            }
        }

        private bool UseAesZipEncryption
        {
            get
            {
                return Properties.Settings.Default.ZipUseAesEncryption;
            }
        }

        private bool ValidateZipAfterCreate
        {
            get
            {
                return Properties.Settings.Default.ZipValidateAfterCreate;
            }
        }

        private bool NormalizeZipEntryNameToAscii
        {
            get
            {
                return Properties.Settings.Default.ZipNormalizeEntryNameToAscii;
            }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // UI スレッドで WinForms コントロールを生成し、BeginInvoke による
            // スレッドマーシャリングのアンカーとして使用する。
            // SynchronizationContext.Post は VSTO 環境では配送が保証されないため使わない。
            this.marshalingControl = new Control();
            this.marshalingControl.CreateControl();

            this.googleDriveService = new GoogleDriveService();
            this.Application.ItemSend += this.Application_ItemSend;
            this.CreateSettingsCommandBarButton();
            this.TryDisableOutlookAttachmentLimit();

            WriteLog("=== Startup: アドイン起動 HasClientConfig=" + this.googleDriveService.HasRequiredClientConfiguration()
                + " RootFolderId=" + (string.IsNullOrWhiteSpace(Properties.Settings.Default.GoogleDriveRootFolderId) ? "(未設定)" : "設定済み") + " ===");

            if (!this.googleDriveService.HasRequiredClientConfiguration() || string.IsNullOrWhiteSpace(Properties.Settings.Default.GoogleDriveRootFolderId))
            {
                this.ShowSettingsForm();
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.ItemSend -= this.Application_ItemSend;
            this.RemoveSettingsCommandBarButton();
            this.RestoreOutlookAttachmentLimit();

            if (this.marshalingControl != null)
            {
                this.marshalingControl.Dispose();
                this.marshalingControl = null;
            }
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            Outlook.MailItem mailItem = item as Outlook.MailItem;
            if (mailItem == null)
            {
                return;
            }

            WriteLog("ItemSend: 発火 Subject=" + (mailItem.Subject ?? "(件名なし)"));

            if (this.ShouldBypassProcessing(mailItem))
            {
                WriteLog("ItemSend: SkipFlag=True → 送信スキップ（アップロード完了後の再送信）");
                this.ClearSkipProcessingFlag(mailItem);
                return;
            }

            if (this.IsCurrentlyProcessing(mailItem))
            {
                WriteLog("ItemSend: 処理中フラグあり → 待機ダイアログ表示");
                cancel = true;
                MessageBox.Show(
                    "このメッセージは現在処理中です。完了までお待ちください。",
                    "処理中",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            IList<AttachmentUploadInfo> uploadableAttachments;
            try
            {
                uploadableAttachments = this.GetUploadableAttachments(mailItem);
            }
            catch (Exception ex)
            {
                WriteLog("ItemSend: GetUploadableAttachments 例外 → " + ex.Message);
                cancel = true;
                MessageBox.Show(
                    "添付ファイルの確認中にエラーが発生しました。送信を中止します。" + Environment.NewLine + ex.Message,
                    "送信前処理エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            WriteLog("ItemSend: アップロード対象件数=" + uploadableAttachments.Count);

            if (uploadableAttachments.Count == 0)
            {
                WriteLog("ItemSend: 添付なし → そのまま送信");
                return;
            }

            if (this.ContainsPassThroughArchiveAttachment(uploadableAttachments))
            {
                WriteLog("ItemSend: .zip/.7z 添付あり → 規定によりそのまま送信");
                return;
            }

            SendProcessingMode mode = this.PromptSendProcessingMode();
            if (mode != SendProcessingMode.GoogleDrive && mode != SendProcessingMode.EncryptedZipAttachment)
            {
                cancel = true;
                WriteLog("ItemSend: ユーザーが送信を中止");
                return;
            }

            if (mode == SendProcessingMode.GoogleDrive && !this.googleDriveService.HasRequiredClientConfiguration())
            {
                WriteLog("ItemSend: client_id/client_secret 未設定 → 設定画面");
                cancel = true;
                MessageBox.Show(
                    "Google Drive の client_id / client_secret が設定されていません。App.config を確認してください。",
                    "Google Drive 設定エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                this.ShowSettingsForm();
                return;
            }

            cancel = true;
            WriteLog("ItemSend: cancel=True、処理開始 Mode=" + mode);

            MailSendOperationContext operationContext;
            try
            {
                operationContext = this.CreateOperationContext(mailItem, uploadableAttachments, mode);
            }
            catch (Exception ex)
            {
                WriteLog("ItemSend: CreateOperationContext 例外 → " + ex.Message);
                MessageBox.Show(
                    "添付ファイルの一時処理中にエラーが発生しました。" + Environment.NewLine + ex.Message,
                    "送信前処理エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            lock (this.processingLock)
            {
                this.processingOperations[operationContext.OperationId] = operationContext;
            }

            try
            {
                this.SetProcessingFlag(mailItem, true);
                WriteLog("ItemSend: ProcessingFlag 設定完了");
            }
            catch (Exception ex)
            {
                WriteLog("ItemSend: SetProcessingFlag 例外（継続） → " + ex.Message);
            }

            if (mode == SendProcessingMode.EncryptedZipAttachment)
            {
                this.ExecuteEncryptedZipAttachmentSend(operationContext);
                return;
            }

            WriteLog("ItemSend: アップロードタスク開始 OperationId=" + operationContext.OperationId);
            UploadProgressForm progressForm = new UploadProgressForm();
            IntPtr ownerHandle = GetForegroundWindow();
            progressForm.Show(ownerHandle != IntPtr.Zero ? (IWin32Window)new Win32Window(ownerHandle) : null);
            operationContext.ProgressForm = progressForm;

            Task.Run(() => this.ProcessMailSendOperationAsync(operationContext));
            WriteLog("ItemSend: Task スケジュール完了");
        }

        private async Task ProcessMailSendOperationAsync(MailSendOperationContext operationContext)
        {
            WriteLog("ProcessMailSendOperationAsync: 開始");

            // UI スレッドセーフな進捗レポーターを構築する
            Action<string> reportProgress = message =>
            {
                Control ctrl2 = this.marshalingControl;
                UploadProgressForm form = operationContext.ProgressForm;
                if (form == null || ctrl2 == null || ctrl2.IsDisposed || !ctrl2.IsHandleCreated) return;
                ctrl2.BeginInvoke(new Action(() => { if (!form.IsDisposed) form.UpdateStatus(message); }));
            };

            DriveUploadResult uploadResult = null;
            Exception uploadException = null;

            try
            {
                uploadResult = await this.googleDriveService.UploadMailAttachmentsAsync(
                    operationContext.Subject,
                    Properties.Settings.Default.GoogleDriveRootFolderId,
                    operationContext.Attachments,
                    CancellationToken.None,
                    reportProgress).ConfigureAwait(false);

                WriteLog("ProcessMailSendOperationAsync: アップロード完了 FolderUrl=" + uploadResult.FolderWebViewLink);
            }
            catch (Exception ex)
            {
                uploadException = ex;
                string innerMessage = ex.InnerException != null ? " | Inner: " + ex.InnerException.Message : string.Empty;
                WriteLog("ProcessMailSendOperationAsync: 例外 " + ex.GetType().Name + " → " + ex.Message + innerMessage);
            }

            Control ctrl = this.marshalingControl;
            if (ctrl == null || ctrl.IsDisposed || !ctrl.IsHandleCreated)
            {
                WriteLog("ProcessMailSendOperationAsync: marshalingControl が無効。UI マーシャリング不可。");
                return;
            }

            if (uploadException == null)
            {
                ctrl.BeginInvoke(new Action(() => this.CompleteMailSendSuccess(operationContext, uploadResult)));
            }
            else
            {
                Exception exToDeliver = uploadException;
                ctrl.BeginInvoke(new Action(() => this.HandleMailSendFailure(operationContext, exToDeliver)));
            }
        }

        private void CompleteMailSendSuccess(MailSendOperationContext operationContext, DriveUploadResult uploadResult)
        {
            WriteLog("CompleteMailSendSuccess: 開始 FolderUrl=" + uploadResult.FolderWebViewLink);
            try
            {
                operationContext.ProgressForm?.UpdateStatus("本文に URL を挿入中...");
                this.InsertFolderLinkAtTop(operationContext.MailItem, uploadResult.FolderWebViewLink);
                WriteLog("CompleteMailSendSuccess: 本文 URL 挿入完了");

                operationContext.ProgressForm?.UpdateStatus("添付ファイルを削除中...");
                this.RemoveUploadedAttachments(operationContext.MailItem, operationContext.OriginalAttachments ?? operationContext.Attachments);
                WriteLog("CompleteMailSendSuccess: 添付ファイル削除完了");

                this.SetProcessingFlag(operationContext.MailItem, false);
                this.MarkForSkipProcessing(operationContext.MailItem);
                operationContext.MailItem.Save();

                operationContext.ProgressForm?.UpdateStatus("メールを送信中...");
                WriteLog("CompleteMailSendSuccess: Save 完了、Send 実行");
                operationContext.MailItem.Send();
                WriteLog("CompleteMailSendSuccess: Send 完了");

                if (this.IsPasswordZipWorkflowEnabled && !string.IsNullOrWhiteSpace(operationContext.ZipPassword))
                {
                    this.CreatePasswordNotificationDraft(operationContext);
                }
            }
            catch (Exception ex)
            {
                WriteLog("CompleteMailSendSuccess: 例外 → " + ex.Message);
                this.CloseProgressForm(operationContext);
                MessageBox.Show(
                    "Google Drive へのアップロードは完了しましたが、メールの最終送信処理に失敗しました。" + Environment.NewLine + ex.Message,
                    "送信エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                this.CloseProgressForm(operationContext);
                this.ReleaseOperation(operationContext);
            }
        }

        private void HandleMailSendFailure(MailSendOperationContext operationContext, Exception ex)
        {
            WriteLog("HandleMailSendFailure: アップロード失敗 → " + ex.Message);
            this.CloseProgressForm(operationContext);

            DialogResult dialogResult = MessageBox.Show(
                "Google Drive へのアップロードに失敗しました。" + Environment.NewLine +
                ex.Message + Environment.NewLine + Environment.NewLine +
                "[はい] を選択すると添付ファイルをそのまま送信します。[いいえ] を選択すると送信を中止します。",
                "アップロード失敗",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    string validationMessage;
                    if (!this.CanSendAsAttachmentWithinConfiguredLimit(operationContext.MailItem, out validationMessage))
                    {
                        WriteLog("HandleMailSendFailure: 添付サイズ超過のため送信中止");
                        MessageBox.Show(
                            validationMessage,
                            "添付サイズ超過",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                    else
                    {
                        this.SetProcessingFlag(operationContext.MailItem, false);
                        this.MarkForSkipProcessing(operationContext.MailItem);
                        operationContext.MailItem.Save();
                        operationContext.MailItem.Send();
                    }
                }
                catch (Exception sendEx)
                {
                    MessageBox.Show(
                        "添付ファイルをそのまま送信することもできませんでした。" + Environment.NewLine + sendEx.Message,
                        "送信失敗",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }

            this.ReleaseOperation(operationContext);
        }

        private MailSendOperationContext CreateOperationContext(Outlook.MailItem mailItem, IList<AttachmentUploadInfo> uploadableAttachments, SendProcessingMode mode)
        {
            string operationId = Guid.NewGuid().ToString("N");
            string temporaryRootDirectory = Path.Combine(Path.GetTempPath(), "UNF_MailFile_GDUpload", operationId);
            Directory.CreateDirectory(temporaryRootDirectory);

            mailItem.Save();

            string originalSubject = mailItem.Subject ?? string.Empty;
            string originalTo = mailItem.To ?? string.Empty;
            string originalCc = mailItem.CC ?? string.Empty;
            string originalBcc = mailItem.BCC ?? string.Empty;
            Outlook.OlBodyFormat originalBodyFormat = mailItem.BodyFormat;

            List<AttachmentUploadInfo> exportedOriginalAttachments = new List<AttachmentUploadInfo>();

            foreach (AttachmentUploadInfo attachment in uploadableAttachments)
            {
                string sanitizedFileName = this.BuildUniqueSafeFileName(temporaryRootDirectory, attachment.FileName);
                string temporaryFilePath = Path.Combine(temporaryRootDirectory, sanitizedFileName);

                Outlook.Attachment outlookAttachment = null;
                try
                {
                    outlookAttachment = mailItem.Attachments[attachment.OriginalIndex];
                    outlookAttachment.SaveAsFile(temporaryFilePath);
                }
                finally
                {
                    if (outlookAttachment != null)
                    {
                        Marshal.ReleaseComObject(outlookAttachment);
                    }
                }

                exportedOriginalAttachments.Add(new AttachmentUploadInfo
                {
                    OriginalIndex = attachment.OriginalIndex,
                    FileName = attachment.FileName,
                    TemporaryFilePath = temporaryFilePath
                });
            }

            string zipPassword = string.Empty;
            string zipFileName = string.Empty;
            List<AttachmentUploadInfo> processingAttachments = exportedOriginalAttachments;

            if (mode == SendProcessingMode.EncryptedZipAttachment)
            {
                zipPassword = this.GenerateZipPassword();
                zipFileName = this.BuildZipFileName(mailItem.Subject);
                string zipFilePath = Path.Combine(temporaryRootDirectory, zipFileName);

                this.CreatePasswordProtectedZip(exportedOriginalAttachments, zipFilePath, zipPassword);

                processingAttachments = new List<AttachmentUploadInfo>
                {
                    new AttachmentUploadInfo
                    {
                        OriginalIndex = 0,
                        FileName = zipFileName,
                        TemporaryFilePath = zipFilePath
                    }
                };
            }

            return new MailSendOperationContext
            {
                OperationId = operationId,
                Mode = mode,
                MailItem = mailItem,
                Subject = originalSubject,
                Attachments = processingAttachments,
                OriginalAttachments = uploadableAttachments.ToList(),
                TemporaryRootDirectory = temporaryRootDirectory,
                ZipPassword = zipPassword,
                ZipFileName = zipFileName,
                OriginalTo = originalTo,
                OriginalCc = originalCc,
                OriginalBcc = originalBcc,
                OriginalBodyFormat = originalBodyFormat
            };
        }

        private void ReleaseOperation(MailSendOperationContext operationContext)
        {
            lock (this.processingLock)
            {
                this.processingOperations.Remove(operationContext.OperationId);
            }

            if (!string.IsNullOrWhiteSpace(operationContext.TemporaryRootDirectory) && Directory.Exists(operationContext.TemporaryRootDirectory))
            {
                try
                {
                    Directory.Delete(operationContext.TemporaryRootDirectory, true);
                }
                catch
                {
                }
            }

            if (operationContext.MailItem != null)
            {
                try
                {
                    this.SetProcessingFlag(operationContext.MailItem, false);
                }
                catch (Exception ex)
                {
                    WriteLog("ReleaseOperation: SetProcessingFlag 例外（無視） → " + ex.Message);
                }
            }

            operationContext.MailItem = null;
        }

        private IList<AttachmentUploadInfo> GetUploadableAttachments(Outlook.MailItem mailItem)
        {
            List<AttachmentUploadInfo> attachments = new List<AttachmentUploadInfo>();
            Outlook.Attachments outlookAttachments = null;

            string htmlBody;
            try
            {
                htmlBody = mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML
                    ? (mailItem.HTMLBody ?? string.Empty)
                    : string.Empty;
            }
            catch
            {
                htmlBody = string.Empty;
            }

            try
            {
                outlookAttachments = mailItem.Attachments;
                int count = outlookAttachments.Count;
                WriteLog("GetUploadableAttachments: Attachments.Count=" + count);

                for (int index = 1; index <= count; index++)
                {
                    Outlook.Attachment attachment = null;
                    try
                    {
                        attachment = outlookAttachments[index];

                        string fileName = attachment.FileName;
                        Outlook.OlAttachmentType attachType = attachment.Type;

                        if (attachType != Outlook.OlAttachmentType.olByValue)
                        {
                            WriteLog("GetUploadableAttachments: [" + index + "] \"" + fileName + "\" Type=" + attachType + " → スキップ(非ByValue)");
                            continue;
                        }

                        bool isInline = this.IsInlineAttachment(attachment, htmlBody);
                        WriteLog("GetUploadableAttachments: [" + index + "] \"" + fileName + "\" IsInline=" + isInline + " → " + (isInline ? "スキップ(インライン)" : "アップロード対象"));

                        if (isInline)
                        {
                            continue;
                        }

                        attachments.Add(new AttachmentUploadInfo
                        {
                            OriginalIndex = index,
                            FileName = fileName
                        });
                    }
                    catch (Exception ex)
                    {
                        WriteLog("GetUploadableAttachments: [" + index + "] 例外 → " + ex.Message);
                    }
                    finally
                    {
                        if (attachment != null)
                        {
                            Marshal.ReleaseComObject(attachment);
                        }
                    }
                }
            }
            finally
            {
                if (outlookAttachments != null)
                {
                    Marshal.ReleaseComObject(outlookAttachments);
                }
            }

            return attachments;
        }

        private bool IsInlineAttachment(Outlook.Attachment attachment, string htmlBody)
        {
            Outlook.PropertyAccessor propertyAccessor = null;
            try
            {
                propertyAccessor = attachment.PropertyAccessor;

                // PR_ATTACHMENT_HIDDEN が true → 署名などの非表示添付として確定
                try
                {
                    object hiddenValue = propertyAccessor.GetProperty(AttachmentHiddenPropertyTag);
                    if (hiddenValue is bool hidden && hidden)
                    {
                        return true;
                    }
                }
                catch
                {
                }

                // PR_ATTACH_FLAGS の ATT_MHTML_REF ビットが立っている → MHTML インライン参照として確定
                try
                {
                    object flagsValue = propertyAccessor.GetProperty(AttachmentFlagsPropertyTag);
                    if (flagsValue != null && (Convert.ToInt32(flagsValue) & AttachmentFlagMhtmlRef) != 0)
                    {
                        return true;
                    }
                }
                catch
                {
                }

                // Content-ID が存在し HTML 本文で cid: として参照されている場合のみインラインと判定する。
                // MAPI の Content-ID は <image001@...> のように角括弧付きで格納される場合があるため正規化する。
                try
                {
                    object contentIdValue = propertyAccessor.GetProperty(AttachmentContentIdPropertyTag);
                    string contentId = contentIdValue as string;
                    if (!string.IsNullOrWhiteSpace(contentId) && !string.IsNullOrWhiteSpace(htmlBody))
                    {
                        string normalizedContentId = contentId.Trim('<', '>');
                        if (!string.IsNullOrWhiteSpace(normalizedContentId))
                        {
                            return htmlBody.IndexOf("cid:" + normalizedContentId, StringComparison.OrdinalIgnoreCase) >= 0;
                        }
                    }
                }
                catch
                {
                }

                return false;
            }
            catch
            {
                // PropertyAccessor の取得を含むすべての処理が失敗した場合は非インライン扱いにする
                return false;
            }
            finally
            {
                if (propertyAccessor != null)
                {
                    Marshal.ReleaseComObject(propertyAccessor);
                }
            }
        }

        private void InsertFolderLinkAtTop(Outlook.MailItem mailItem, string folderUrl)
        {
            string normalizedUrl = string.IsNullOrWhiteSpace(folderUrl)
                ? string.Empty
                : folderUrl.Trim();
            string expirationDateText = DateTime.Now.Date.AddDays(6).ToString("yyyy/MM/dd");

            const string separator = "************************************************";
            string plainPrefix = separator + Environment.NewLine +
                                  "本メールの添付ファイルは、" + Environment.NewLine +
                                  "下記URLより確認・ダウンロードをお願いいたします。" + Environment.NewLine +
                                  normalizedUrl + Environment.NewLine +
                                  "※上記URLの有効期限は、" + expirationDateText + " です。（送信日含めて7日間です。）" + Environment.NewLine +
                                  separator + Environment.NewLine + Environment.NewLine;

            string htmlPrefix = "<div style=\"margin-bottom:12px;font-family:monospace;\">" +
                                 separator + "<br />" +
                                 "本メールの添付ファイルは、<br />" +
                                 "下記URLより確認・ダウンロードをお願いいたします。<br />" +
                                 "<a href=\"" + System.Security.SecurityElement.Escape(normalizedUrl) + "\">" +
                                 System.Security.SecurityElement.Escape(normalizedUrl) + "</a><br />" +
                                 "※上記URLの有効期限は、" + System.Security.SecurityElement.Escape(expirationDateText) + " です。（送信日含めて7日間です。）<br />" +
                                 separator +
                                 "</div>";

            if (mailItem.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
            {
                string currentHtml = mailItem.HTMLBody ?? string.Empty;
                if (string.IsNullOrWhiteSpace(currentHtml))
                {
                    mailItem.HTMLBody = "<html><body>" + htmlPrefix + "</body></html>";
                    return;
                }

                int bodyTagIndex = currentHtml.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                if (bodyTagIndex >= 0)
                {
                    int bodyStartCloseIndex = currentHtml.IndexOf('>', bodyTagIndex);
                    if (bodyStartCloseIndex >= 0)
                    {
                        mailItem.HTMLBody = currentHtml.Insert(bodyStartCloseIndex + 1, htmlPrefix);
                        return;
                    }
                }

                mailItem.HTMLBody = htmlPrefix + currentHtml;
                return;
            }

            mailItem.Body = plainPrefix + (mailItem.Body ?? string.Empty);
        }

        private void RemoveUploadedAttachments(Outlook.MailItem mailItem, IList<AttachmentUploadInfo> uploadedAttachments)
        {
            // アップロード対象ファイル名のセットを構築する。
            // ファイル名ベースで照合することで、非同期処理中にインデックスがずれても正しく削除できる。
            HashSet<string> fileNamesToRemove = new HashSet<string>(
                uploadedAttachments.Select(a => a.FileName),
                StringComparer.OrdinalIgnoreCase);

            Outlook.Attachments outlookAttachments = null;
            try
            {
                outlookAttachments = mailItem.Attachments;
                int count = outlookAttachments.Count;

                // 後ろから削除することで削除によるインデックスのずれを防ぐ
                for (int i = count; i >= 1; i--)
                {
                    Outlook.Attachment attachment = null;
                    try
                    {
                        attachment = outlookAttachments[i];
                        if (fileNamesToRemove.Contains(attachment.FileName))
                        {
                            attachment.Delete();
                        }
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (attachment != null)
                        {
                            Marshal.ReleaseComObject(attachment);
                        }
                    }
                }
            }
            finally
            {
                if (outlookAttachments != null)
                {
                    Marshal.ReleaseComObject(outlookAttachments);
                }
            }
        }

        private string GenerateZipPassword()
        {
            Random random = new Random(Guid.NewGuid().GetHashCode());
            int number = random.Next(0, 100000000);
            return ZipPasswordPrefix + number.ToString("D" + ZipPasswordDigits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        }

        private string BuildZipFileName(string subject)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            string safeSubject = string.IsNullOrWhiteSpace(subject) ? "NoSubject" : subject.Trim();

            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                safeSubject = safeSubject.Replace(invalidCharacter, '_');
            }

            // Outlook/Windows 互換のため追加で置換
            safeSubject = safeSubject.Replace('[', '_')
                                     .Replace(']', '_')
                                     .Replace('#', '_');

            if (safeSubject.Length > 120)
            {
                safeSubject = safeSubject.Substring(0, 120);
            }

            if (string.IsNullOrWhiteSpace(safeSubject))
            {
                safeSubject = "NoSubject";
            }

            return timestamp + "_" + safeSubject + ".zip";
        }

        private void CreatePasswordProtectedZip(IList<AttachmentUploadInfo> sourceAttachments, string zipFilePath, string zipPassword)
        {
            using (FileStream fileStream = File.Create(zipFilePath))
            using (ZipOutputStream zipStream = new ZipOutputStream(fileStream))
            {
                zipStream.SetLevel(9);
                zipStream.Password = zipPassword;
                zipStream.UseZip64 = UseZip64.Off;

                foreach (AttachmentUploadInfo attachment in sourceAttachments)
                {
                    string rawEntryName = Path.GetFileName(attachment.TemporaryFilePath);
                    string entryName = this.NormalizeZipEntryNameToAscii
                        ? this.NormalizeEntryNameForLegacyTools(rawEntryName)
                        : ZipEntry.CleanName(rawEntryName);

                    ZipEntry entry = new ZipEntry(entryName)
                    {
                        DateTime = DateTime.Now,
                        IsUnicodeText = !this.NormalizeZipEntryNameToAscii
                    };

                    if (this.UseAesZipEncryption)
                    {
                        entry.AESKeySize = 256;
                    }

                    zipStream.PutNextEntry(entry);
                    using (FileStream inputStream = File.OpenRead(attachment.TemporaryFilePath))
                    {
                        inputStream.CopyTo(zipStream);
                    }
                    zipStream.CloseEntry();
                }

                zipStream.IsStreamOwner = true;
                zipStream.Finish();
            }

            if (this.ValidateZipAfterCreate)
            {
                this.VerifyPasswordZip(zipFilePath, zipPassword);
            }
        }

        private string NormalizeEntryNameForLegacyTools(string fileName)
        {
            string candidate = ZipEntry.CleanName(fileName ?? string.Empty);
            if (string.IsNullOrWhiteSpace(candidate))
            {
                return "attachment.bin";
            }

            StringBuilder builder = new StringBuilder(candidate.Length);
            for (int i = 0; i < candidate.Length; i++)
            {
                char c = candidate[i];
                bool keep =
                    (c >= 'A' && c <= 'Z') ||
                    (c >= 'a' && c <= 'z') ||
                    (c >= '0' && c <= '9') ||
                    c == '.' || c == '_' || c == '-' || c == '(' || c == ')' || c == ' ';
                builder.Append(keep ? c : '_');
            }

            string normalized = builder.ToString().Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                normalized = "attachment.bin";
            }

            return normalized;
        }

        private void VerifyPasswordZip(string zipFilePath, string zipPassword)
        {
            try
            {
                using (FileStream zipStream = File.OpenRead(zipFilePath))
                using (ZipFile zipFile = new ZipFile(zipStream))
                {
                    zipFile.Password = zipPassword;

                    foreach (ZipEntry entry in zipFile)
                    {
                        if (entry == null || !entry.IsFile)
                        {
                            continue;
                        }

                        using (Stream entryStream = zipFile.GetInputStream(entry))
                        {
                            byte[] buffer = new byte[1];
                            entryStream.Read(buffer, 0, 1);
                        }
                    }
                }

                WriteLog("VerifyPasswordZip: 検証成功 Path=" + zipFilePath);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("作成した Zip ファイルの整合性検証に失敗しました。", ex);
            }
        }

        private void CreatePasswordNotificationDraft(MailSendOperationContext operationContext)
        {
            Outlook.MailItem draftMail = null;
            try
            {
                draftMail = this.Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                if (draftMail == null)
                {
                    return;
                }

                draftMail.Subject = "【PW通知】" + (operationContext.Subject ?? string.Empty);
                draftMail.To = this.NormalizeRecipientAddressList(operationContext.OriginalTo);
                draftMail.CC = this.NormalizeRecipientAddressList(operationContext.OriginalCc);
                draftMail.BCC = this.NormalizeRecipientAddressList(operationContext.OriginalBcc);

                string passwordText = operationContext.ZipPassword ?? string.Empty;

                if (operationContext.OriginalBodyFormat == Outlook.OlBodyFormat.olFormatHTML)
                {
                    string existingHtml = draftMail.HTMLBody ?? string.Empty;
                    string htmlTop =
                        "<div>先程送信しましたメール添付ファイルのPWをお知らせいたします。<br />" +
                        System.Security.SecurityElement.Escape(passwordText) + "<br />" +
                        "※本Zipファイルは、Googleドライブより添付ファイルをダウンロードした後に必要となります。" +
                        "</div><br />";
                    draftMail.HTMLBody = htmlTop + existingHtml;
                }
                else
                {
                    string bodyTop =
                        "先程送信しましたメール添付ファイルのPWをお知らせいたします。" + Environment.NewLine +
                        passwordText + Environment.NewLine +
                        "※本Zipファイルは、Googleドライブより添付ファイルをダウンロードした後に必要となります。" +
                        Environment.NewLine + Environment.NewLine;

                    string existingBody = draftMail.Body ?? string.Empty;
                    draftMail.Body = bodyTop + existingBody;
                }

                draftMail.Save();
                WriteLog("CreatePasswordNotificationDraft: 下書き保存完了 Subject=" + draftMail.Subject);
            }
            catch (Exception ex)
            {
                WriteLog("CreatePasswordNotificationDraft: 例外 → " + ex.Message + " | " + ex.GetType().Name);
                MessageBox.Show(
                    "PW通知メールの下書き作成に失敗しました。" + Environment.NewLine + ex.Message,
                    "下書き作成エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                if (draftMail != null)
                {
                    Marshal.ReleaseComObject(draftMail);
                }
            }
        }

        private string NormalizeRecipientAddressList(string addressList)
        {
            if (string.IsNullOrWhiteSpace(addressList))
            {
                return string.Empty;
            }

            string[] tokens = addressList.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> normalizedTokens = new List<string>();

            foreach (string token in tokens)
            {
                string value = token.Trim();
                if (value.Length == 0)
                {
                    continue;
                }

                if ((value.StartsWith("'", StringComparison.Ordinal) && value.EndsWith("'", StringComparison.Ordinal)) ||
                    (value.StartsWith("\"", StringComparison.Ordinal) && value.EndsWith("\"", StringComparison.Ordinal)))
                {
                    if (value.Length > 1)
                    {
                        value = value.Substring(1, value.Length - 2).Trim();
                    }
                }

                normalizedTokens.Add(value);
            }

            return string.Join("; ", normalizedTokens);
        }

        private bool CanSendAsAttachmentWithinConfiguredLimit(Outlook.MailItem mailItem, out string validationMessage)
        {
            validationMessage = string.Empty;

            long configuredLimitKb;
            if (!this.TryGetConfiguredMaximumAttachmentSizeKb(out configuredLimitKb) || configuredLimitKb <= 0)
            {
                WriteLog("CanSendAsAttachmentWithinConfiguredLimit: MaximumAttachmentSize 未設定または無制限のためチェック不要");
                return true;
            }

            long totalAttachmentBytes = this.CalculateTotalAttachmentSizeBytes(mailItem);
            long limitBytes = configuredLimitKb * 1024L;
            if (totalAttachmentBytes <= limitBytes)
            {
                return true;
            }

            validationMessage =
                "添付ファイルをそのまま送信しようとしましたが、" + Environment.NewLine +
                "添付合計サイズが Outlook の上限を超えています。" + Environment.NewLine + Environment.NewLine +
                "上限: " + this.FormatSize(limitBytes) + Environment.NewLine +
                "現在: " + this.FormatSize(totalAttachmentBytes) + Environment.NewLine + Environment.NewLine +
                "Google Drive のアップロード再実行、または添付ファイルを減らして再送してください。";

            return false;
        }

        private bool TryGetConfiguredMaximumAttachmentSizeKb(out long sizeKb)
        {
            sizeKb = 0;
            try
            {
                if (this.hadMaximumAttachmentSizeValue && this.originalMaximumAttachmentSizeValue != null)
                {
                    sizeKb = Convert.ToInt64(this.originalMaximumAttachmentSizeValue);
                    return true;
                }

                if (string.IsNullOrWhiteSpace(this.outlookPreferencesRegistryPath))
                {
                    return false;
                }

                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(this.outlookPreferencesRegistryPath, false))
                {
                    object currentValue = key?.GetValue(OutlookMaximumAttachmentSizeValueName, null);
                    if (currentValue == null)
                    {
                        return false;
                    }

                    sizeKb = Convert.ToInt64(currentValue);
                    return true;
                }
            }
            catch
            {
                sizeKb = 0;
                return false;
            }
        }

        private long CalculateTotalAttachmentSizeBytes(Outlook.MailItem mailItem)
        {
            long totalBytes = 0;
            Outlook.Attachments attachments = null;
            try
            {
                attachments = mailItem.Attachments;
                for (int index = 1; index <= attachments.Count; index++)
                {
                    Outlook.Attachment attachment = null;
                    try
                    {
                        attachment = attachments[index];
                        totalBytes += attachment.Size;
                    }
                    finally
                    {
                        if (attachment != null) Marshal.ReleaseComObject(attachment);
                    }
                }
            }
            finally
            {
                if (attachments != null) Marshal.ReleaseComObject(attachments);
            }

            return totalBytes;
        }

        private string FormatSize(long bytes)
        {
            const double kb = 1024d;
            const double mb = kb * 1024d;
            const double gb = mb * 1024d;

            if (bytes >= gb) return (bytes / gb).ToString("0.00") + " GB";
            if (bytes >= mb) return (bytes / mb).ToString("0.00") + " MB";
            if (bytes >= kb) return (bytes / kb).ToString("0.00") + " KB";
            return bytes + " bytes";
        }

        private string BuildUniqueSafeFileName(string directoryPath, string originalFileName)
        {
            string candidateFileName = string.IsNullOrWhiteSpace(originalFileName) ? "attachment.bin" : originalFileName;
            foreach (char invalidCharacter in Path.GetInvalidFileNameChars())
            {
                candidateFileName = candidateFileName.Replace(invalidCharacter, '_');
            }

            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(candidateFileName);
            string extension = Path.GetExtension(candidateFileName);
            string fullPath = Path.Combine(directoryPath, candidateFileName);
            int sequence = 1;
            while (File.Exists(fullPath))
            {
                candidateFileName = fileNameWithoutExtension + "_" + sequence.ToString() + extension;
                fullPath = Path.Combine(directoryPath, candidateFileName);
                sequence++;
            }

            return candidateFileName;
        }

        private bool ShouldBypassProcessing(Outlook.MailItem mailItem)
        {
            Outlook.UserProperty userProperty = null;
            try
            {
                userProperty = mailItem.UserProperties.Find(SkipProcessingUserPropertyName);
                return userProperty != null && string.Equals(Convert.ToString(userProperty.Value), bool.TrueString, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                if (userProperty != null) Marshal.ReleaseComObject(userProperty);
            }
        }

        private bool IsCurrentlyProcessing(Outlook.MailItem mailItem)
        {
            Outlook.UserProperty userProperty = null;
            try
            {
                userProperty = mailItem.UserProperties.Find(ProcessingUserPropertyName);
                return userProperty != null && string.Equals(Convert.ToString(userProperty.Value), bool.TrueString, StringComparison.OrdinalIgnoreCase);
            }
            finally
            {
                if (userProperty != null) Marshal.ReleaseComObject(userProperty);
            }
        }

        private void MarkForSkipProcessing(Outlook.MailItem mailItem)
        {
            Outlook.UserProperty userProperty = null;
            try
            {
                userProperty = mailItem.UserProperties.Find(SkipProcessingUserPropertyName) ??
                               mailItem.UserProperties.Add(SkipProcessingUserPropertyName, Outlook.OlUserPropertyType.olText, false, Type.Missing);
                userProperty.Value = bool.TrueString;
            }
            finally
            {
                if (userProperty != null) Marshal.ReleaseComObject(userProperty);
            }
        }

        private void ClearSkipProcessingFlag(Outlook.MailItem mailItem)
        {
            Outlook.UserProperty userProperty = null;
            try
            {
                userProperty = mailItem.UserProperties.Find(SkipProcessingUserPropertyName);
                if (userProperty != null)
                {
                    userProperty.Value = string.Empty;
                }
            }
            finally
            {
                if (userProperty != null) Marshal.ReleaseComObject(userProperty);
            }
        }

        private void SetProcessingFlag(Outlook.MailItem mailItem, bool isProcessing)
        {
            Outlook.UserProperty userProperty = null;
            try
            {
                userProperty = mailItem.UserProperties.Find(ProcessingUserPropertyName) ??
                               mailItem.UserProperties.Add(ProcessingUserPropertyName, Outlook.OlUserPropertyType.olText, false, Type.Missing);
                userProperty.Value = isProcessing ? bool.TrueString : string.Empty;
            }
            finally
            {
                if (userProperty != null) Marshal.ReleaseComObject(userProperty);
            }
        }

        private void CreateSettingsCommandBarButton()
        {
            Outlook.Explorer activeExplorer = null;
            CommandBars commandBars = null;
            CommandBar targetCommandBar = null;
            try
            {
                activeExplorer = this.Application.ActiveExplorer();
                commandBars = activeExplorer?.CommandBars;
                targetCommandBar = commandBars?["Standard"] ?? commandBars?["Menu Bar"];
                if (targetCommandBar == null)
                {
                    return;
                }

                this.settingsCommandBarButton = targetCommandBar.Controls.Add(MsoControlType.msoControlButton, Temporary: true) as CommandBarButton;
                if (this.settingsCommandBarButton == null)
                {
                    return;
                }

                this.settingsCommandBarButton.Caption = "Google Drive 設定";
                this.settingsCommandBarButton.Style = MsoButtonStyle.msoButtonCaption;
                this.settingsCommandBarButton.Tag = "UNF_MailFile_GDUpload_Settings";
                this.settingsCommandBarButton.Click += this.SettingsCommandBarButton_Click;
            }
            catch
            {
            }
            finally
            {
                if (targetCommandBar != null) Marshal.ReleaseComObject(targetCommandBar);
                if (commandBars != null) Marshal.ReleaseComObject(commandBars);
                if (activeExplorer != null) Marshal.ReleaseComObject(activeExplorer);
            }
        }

        private void RemoveSettingsCommandBarButton()
        {
            if (this.settingsCommandBarButton == null)
            {
                return;
            }

            try
            {
                this.settingsCommandBarButton.Click -= this.SettingsCommandBarButton_Click;
                this.settingsCommandBarButton.Delete(false);
            }
            catch
            {
            }
            finally
            {
                Marshal.ReleaseComObject(this.settingsCommandBarButton);
                this.settingsCommandBarButton = null;
            }
        }

        private void SettingsCommandBarButton_Click(CommandBarButton ctrl, ref bool cancelDefault)
        {
            this.ShowSettingsForm();
        }

        private void ShowSettingsForm()
        {
            using (SettingsForm settingsForm = new SettingsForm(this.googleDriveService))
            {
                settingsForm.StartPosition = FormStartPosition.CenterScreen;
                settingsForm.ShowDialog();
            }
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        private static readonly object logLock = new object();

        private static void WriteLog(string message)
        {
            try
            {
                lock (logLock)
                {
                    string dir = Path.GetDirectoryName(DiagnosticLogPath);
                    if (!Directory.Exists(dir))
                    {
                        Directory.CreateDirectory(dir);
                    }

                    if (File.Exists(DiagnosticLogPath) && new FileInfo(DiagnosticLogPath).Length > 512 * 1024)
                    {
                        File.Delete(DiagnosticLogPath);
                    }

                    using (FileStream stream = new FileStream(DiagnosticLogPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                    using (StreamWriter writer = new StreamWriter(stream, Encoding.UTF8))
                    {
                        writer.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + message);
                    }
                }
            }
            catch
            {
            }
        }

        private void CloseProgressForm(MailSendOperationContext operationContext)
        {
            UploadProgressForm form = operationContext.ProgressForm;
            operationContext.ProgressForm = null;
            if (form == null || form.IsDisposed) return;
            try
            {
                form.Close();
                form.Dispose();
            }
            catch
            {
            }
        }

        private void TryDisableOutlookAttachmentLimit()
        {
            try
            {
                string officeVersion = this.Application != null ? this.Application.Version : "16.0";
                this.outlookPreferencesRegistryPath = @"Software\Microsoft\Office\" + officeVersion + @"\Outlook\Preferences";

                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(this.outlookPreferencesRegistryPath, true))
                {
                    if (key == null)
                    {
                        WriteLog("TryDisableOutlookAttachmentLimit: レジストリキー作成/取得失敗");
                        return;
                    }

                    object currentValue = key.GetValue(OutlookMaximumAttachmentSizeValueName, null);
                    if (currentValue != null)
                    {
                        this.hadMaximumAttachmentSizeValue = true;
                        this.originalMaximumAttachmentSizeValue = currentValue;
                        this.originalMaximumAttachmentSizeValueKind = key.GetValueKind(OutlookMaximumAttachmentSizeValueName);
                    }
                    else
                    {
                        this.hadMaximumAttachmentSizeValue = false;
                        this.originalMaximumAttachmentSizeValue = null;
                        this.originalMaximumAttachmentSizeValueKind = null;
                    }

                    key.SetValue(OutlookMaximumAttachmentSizeValueName, 0, RegistryValueKind.DWord);
                    this.attachmentLimitOverrideApplied = true;
                }

                WriteLog("TryDisableOutlookAttachmentLimit: MaximumAttachmentSize=0 を適用 Path=" + this.outlookPreferencesRegistryPath);
            }
            catch (Exception ex)
            {
                WriteLog("TryDisableOutlookAttachmentLimit: 例外 → " + ex.Message);
            }
        }

        private void RestoreOutlookAttachmentLimit()
        {
            if (!this.attachmentLimitOverrideApplied || string.IsNullOrWhiteSpace(this.outlookPreferencesRegistryPath))
            {
                return;
            }

            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(this.outlookPreferencesRegistryPath, true))
                {
                    if (key == null)
                    {
                        return;
                    }

                    if (this.hadMaximumAttachmentSizeValue)
                    {
                        RegistryValueKind valueKind = this.originalMaximumAttachmentSizeValueKind ?? RegistryValueKind.DWord;
                        object value = this.originalMaximumAttachmentSizeValue ?? 0;
                        key.SetValue(OutlookMaximumAttachmentSizeValueName, value, valueKind);
                        WriteLog("RestoreOutlookAttachmentLimit: MaximumAttachmentSize を元の値に復元");
                    }
                    else
                    {
                        key.DeleteValue(OutlookMaximumAttachmentSizeValueName, false);
                        WriteLog("RestoreOutlookAttachmentLimit: MaximumAttachmentSize を削除（元々未設定）");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("RestoreOutlookAttachmentSize: 例外 → " + ex.Message);
            }
            finally
            {
                this.attachmentLimitOverrideApplied = false;
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private sealed class Win32Window : IWin32Window
        {
            private readonly IntPtr handle;
            internal Win32Window(IntPtr handle) { this.handle = handle; }
            public IntPtr Handle => this.handle;
        }

        private sealed class MailSendOperationContext
        {
            public string OperationId { get; set; }
            public SendProcessingMode Mode { get; set; }
            public Outlook.MailItem MailItem { get; set; }
            public string Subject { get; set; }
            public List<AttachmentUploadInfo> Attachments { get; set; }
            public List<AttachmentUploadInfo> OriginalAttachments { get; set; }
            public string ZipPassword { get; set; }
            public string ZipFileName { get; set; }
            public string OriginalTo { get; set; }
            public string OriginalCc { get; set; }
            public string OriginalBcc { get; set; }
            public Outlook.OlBodyFormat OriginalBodyFormat { get; set; }
            public string TemporaryRootDirectory { get; set; }
            public UploadProgressForm ProgressForm { get; set; }
        }

        private SendProcessingMode PromptSendProcessingMode()
        {
            using (SendModeSelectionForm dialog = new SendModeSelectionForm())
            {
                IntPtr ownerHandle = GetForegroundWindow();
                DialogResult result = ownerHandle != IntPtr.Zero
                    ? dialog.ShowDialog(new Win32Window(ownerHandle))
                    : dialog.ShowDialog();

                if (result == DialogResult.OK && dialog.SelectedMode.HasValue)
                {
                    return dialog.SelectedMode.Value;
                }
            }

            return (SendProcessingMode)(-1);
        }

        private bool ContainsPassThroughArchiveAttachment(IList<AttachmentUploadInfo> attachments)
        {
            foreach (AttachmentUploadInfo attachment in attachments)
            {
                string extension = Path.GetExtension(attachment.FileName) ?? string.Empty;
                if (string.Equals(extension, ".zip", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(extension, ".7z", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private void ExecuteEncryptedZipAttachmentSend(MailSendOperationContext operationContext)
        {
            try
            {
                operationContext.ProgressForm?.UpdateStatus("暗号化Zipを添付中...");
                this.RemoveUploadedAttachments(operationContext.MailItem, operationContext.OriginalAttachments ?? operationContext.Attachments);

                string zipPath = operationContext.Attachments != null && operationContext.Attachments.Count > 0
                    ? operationContext.Attachments[0].TemporaryFilePath
                    : string.Empty;

                if (string.IsNullOrWhiteSpace(zipPath) || !File.Exists(zipPath))
                {
                    throw new FileNotFoundException("作成したZipファイルが見つかりません。", zipPath);
                }

                Outlook.Attachment addedAttachment = null;
                try
                {
                    addedAttachment = operationContext.MailItem.Attachments.Add(
                        zipPath,
                        Outlook.OlAttachmentType.olByValue,
                        Type.Missing,
                        operationContext.ZipFileName ?? Path.GetFileName(zipPath));
                }
                finally
                {
                    if (addedAttachment != null)
                    {
                        Marshal.ReleaseComObject(addedAttachment);
                    }
                }

                string validationMessage;
                if (!this.CanSendAsAttachmentWithinConfiguredLimit(operationContext.MailItem, out validationMessage))
                {
                    MessageBox.Show(validationMessage, "添付サイズ超過", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.ReleaseOperation(operationContext);
                    return;
                }

                this.SetProcessingFlag(operationContext.MailItem, false);
                this.MarkForSkipProcessing(operationContext.MailItem);
                operationContext.MailItem.Save();

                Control ctrl = this.marshalingControl;
                if (ctrl != null && !ctrl.IsDisposed && ctrl.IsHandleCreated)
                {
                    ctrl.BeginInvoke(new Action(() => this.CompleteEncryptedZipAttachmentSend(operationContext)));
                }
                else
                {
                    // フォールバック（通常は通らない）
                    this.CompleteEncryptedZipAttachmentSend(operationContext);
                }
            }
            catch (Exception ex)
            {
                WriteLog("ExecuteEncryptedZipAttachmentSend: 例外 → " + ex.Message);
                MessageBox.Show(
                    "暗号化Zip添付送信に失敗しました。" + Environment.NewLine + ex.Message,
                    "送信エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                this.ReleaseOperation(operationContext);
            }
        }

        private void CompleteEncryptedZipAttachmentSend(MailSendOperationContext operationContext)
        {
            try
            {
                operationContext.MailItem.Send();
                this.CreatePasswordNotificationDraft(operationContext);
            }
            catch (Exception ex)
            {
                WriteLog("CompleteEncryptedZipAttachmentSend: 例外 → " + ex.Message);
                MessageBox.Show(
                    "暗号化Zip添付送信に失敗しました。" + Environment.NewLine + ex.Message,
                    "送信エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                this.ReleaseOperation(operationContext);
            }
        }

        private sealed class SendModeSelectionForm : Form
        {
            private readonly Button btnGoogleDrive;
            private readonly Button btnEncryptedZip;
            private readonly Button btnCancel;

            internal SendProcessingMode? SelectedMode { get; private set; }

            internal SendModeSelectionForm()
            {
                this.Text = "送信方法の選択";
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.StartPosition = FormStartPosition.CenterParent;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowInTaskbar = false;
                this.ClientSize = new System.Drawing.Size(520, 170);

                Label lblMessage = new Label
                {
                    AutoSize = false,
                    Left = 16,
                    Top = 16,
                    Width = 488,
                    Height = 56,
                    Text = "送信方法を選択してください。",
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft
                };

                this.btnGoogleDrive = new Button
                {
                    Left = 16,
                    Top = 86,
                    Width = 238,
                    Height = 34,
                    Text = "Googleドライブに格納して送信"
                };
                this.btnGoogleDrive.Click += this.BtnGoogleDrive_Click;

                this.btnEncryptedZip = new Button
                {
                    Left = 266,
                    Top = 86,
                    Width = 238,
                    Height = 34,
                    Text = "暗号化Zipを作成して添付送信"
                };
                this.btnEncryptedZip.Click += this.BtnEncryptedZip_Click;

                this.btnCancel = new Button
                {
                    Left = 390,
                    Top = 132,
                    Width = 114,
                    Height = 28,
                    Text = "キャンセル",
                    DialogResult = DialogResult.Cancel
                };

                this.CancelButton = this.btnCancel;
                this.Controls.Add(lblMessage);
                this.Controls.Add(this.btnGoogleDrive);
                this.Controls.Add(this.btnEncryptedZip);
                this.Controls.Add(this.btnCancel);
            }

            private void BtnGoogleDrive_Click(object sender, EventArgs e)
            {
                this.SelectedMode = SendProcessingMode.GoogleDrive;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

            private void BtnEncryptedZip_Click(object sender, EventArgs e)
            {
                this.SelectedMode = SendProcessingMode.EncryptedZipAttachment;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }
    }
}
