using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Interop.Outlook;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public class OutlookInterop
    {
        private readonly Application _app;

        public OutlookInterop(Application app)
        {
            _app = app;
        }

        public EmailContext GetCurrentEmailContext()
        {
            var explorer = _app.ActiveExplorer();
            if (explorer?.Selection?.Count > 0)
            {
                var item = explorer.Selection[1];
                if (item is MailItem mail)
                    return ExtractContext(mail, isComposing: false);
            }

            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem composing)
                return ExtractContext(composing, isComposing: true);

            return null;
        }

        public string GetCurrentEntryIdOrTemp()
        {
            var explorer = _app.ActiveExplorer();
            if (explorer?.Selection?.Count > 0)
            {
                var item = explorer.Selection[1];
                if (item is MailItem mail && !string.IsNullOrEmpty(mail.EntryID))
                    return mail.EntryID;
            }

            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem composing)
            {
                if (!string.IsNullOrEmpty(composing.EntryID))
                    return composing.EntryID;
            }

            return "temp-" + Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        public string GetSelectedText()
        {
            try
            {
                var inspector = _app.ActiveInspector();
                if (inspector == null) return null;

                // WordEditor returns a dynamic COM object — no Word interop reference needed
                dynamic doc = inspector.WordEditor;
                if (doc == null) return null;

                dynamic selection = doc.Application.Selection;
                if (selection != null && !string.IsNullOrWhiteSpace((string)selection.Text))
                    return ((string)selection.Text).Trim();
            }
            catch
            {
                // WordEditor may not be available in all contexts
            }
            return null;
        }

        public void ApplyToBody(string newBody)
        {
            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem mail)
            {
                // HTML 본문이 있으면 서명 보존하면서 본문 교체
                if (!string.IsNullOrEmpty(mail.HTMLBody) && mail.HTMLBody.Contains("<body"))
                {
                    InsertHtmlContent(mail, newBody);
                }
                else
                {
                    mail.Body = newBody;
                }
            }
        }

        /// <summary>
        /// Compose/AutoReply 결과를 이메일에 반영.
        /// HTML 서명을 보존하면서 본문을 삽입.
        /// </summary>
        public void ApplyComposeDraft(string body, string subject = null)
        {
            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem composing)
            {
                if (!string.IsNullOrEmpty(subject))
                    composing.Subject = subject;

                InsertHtmlContent(composing, body);
                return;
            }

            // Explorer에서 선택된 메일에 답장
            var explorer = _app.ActiveExplorer();
            if (explorer?.Selection?.Count > 0)
            {
                var item = explorer.Selection[1];
                if (item is MailItem original)
                {
                    var reply = original.Reply();
                    if (!string.IsNullOrEmpty(subject))
                        reply.Subject = subject;

                    InsertHtmlContent(reply, body);
                    reply.Display(false);
                }
            }
        }

        public void InsertReplyDraft(string replyBody)
        {
            ApplyComposeDraft(replyBody);
        }

        /// <summary>
        /// MailItem의 HTMLBody에 내용을 삽입.
        /// 기존 서명과 포맷을 보존하면서 body 태그 바로 뒤에 삽입.
        /// </summary>
        private void InsertHtmlContent(MailItem mail, string plainText)
        {
            var html = mail.HTMLBody ?? "";
            var htmlContent = "<div style='font-family:Calibri,sans-serif;font-size:11pt'>"
                            + System.Net.WebUtility.HtmlEncode(plainText)
                                .Replace("\n", "<br>")
                                .Replace("  ", " &nbsp;")
                            + "</div><br>";

            if (html.Contains("<body"))
            {
                var bodyTag = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                var bodyClose = html.IndexOf(">", bodyTag) + 1;

                // 기존 본문 내용 중 서명 전까지 찾기
                // Outlook 서명은 보통 <div id="Signature"> 또는 <div class="signature"> 등으로 시작
                var sigPatterns = new[] {
                    "id=\"Signature\"", "id='Signature'",
                    "class=\"signature\"", "class='signature'",
                    "id=\"mailSignature\"", "id=\"mail-signature\""
                };

                int sigPos = -1;
                foreach (var pat in sigPatterns)
                {
                    sigPos = html.IndexOf(pat, bodyClose, StringComparison.OrdinalIgnoreCase);
                    if (sigPos > 0)
                    {
                        // div 시작 태그까지 되돌아가기
                        sigPos = html.LastIndexOf("<div", sigPos, sigPos - bodyClose, StringComparison.OrdinalIgnoreCase);
                        break;
                    }
                }

                if (sigPos > 0)
                {
                    // 서명 바로 앞에 삽입
                    mail.HTMLBody = html.Insert(sigPos, htmlContent);
                }
                else
                {
                    // 서명을 못 찾으면 body 태그 바로 뒤에 삽입
                    mail.HTMLBody = html.Insert(bodyClose, htmlContent);
                }
            }
            else
            {
                mail.Body = plainText;
            }
        }

        /// <summary>
        /// 현재 사용자의 표시 이름과 이메일 주소.
        /// </summary>
        public string GetUserDisplayName()
        {
            try
            {
                var ns = _app.GetNamespace("MAPI");
                return ns.CurrentUser?.Name ?? "";
            }
            catch
            {
                return "";
            }
        }

        public string GetUserEmailAddress()
        {
            try
            {
                var ns = _app.GetNamespace("MAPI");
                var account = ns.Accounts[1];
                return account?.SmtpAddress ?? "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Returns the Office UI locale as a language name (e.g., "Korean", "English", "Japanese").
        /// Falls back to system UI culture.
        /// </summary>
        public string GetUserLanguage()
        {
            try
            {
                var lcid = _app.LanguageSettings
                    .get_LanguageID(Microsoft.Office.Core.MsoAppLanguageID.msoLanguageIDUI);
                var culture = CultureInfo.GetCultureInfo(lcid);
                return culture.EnglishName; // e.g., "Korean", "English"
            }
            catch
            {
                return CultureInfo.CurrentUICulture.EnglishName;
            }
        }

        private EmailContext ExtractContext(MailItem mail, bool isComposing)
        {
            return new EmailContext
            {
                Subject = mail.Subject ?? "",
                Body = mail.Body ?? "",
                BodyHtml = mail.HTMLBody ?? "",
                Recipients = GetRecipients(mail),
                SenderEmail = mail.SenderEmailAddress ?? "",
                IsComposing = isComposing
            };
        }

        private string GetRecipients(MailItem mail)
        {
            if (mail.Recipients == null) return "";
            var list = new List<string>();
            foreach (Recipient r in mail.Recipients)
                list.Add(r.Address ?? r.Name ?? "");
            return string.Join("; ", list);
        }
    }
}
