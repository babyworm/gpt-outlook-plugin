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
                mail.Body = newBody;
            }
        }

        public void ApplyToHtmlBody(string newHtmlBody)
        {
            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem mail)
            {
                mail.HTMLBody = newHtmlBody;
            }
        }

        /// <summary>
        /// Compose 결과를 이메일에 반영.
        /// 답장: Reply 생성 후 본문 삽입 (기존 서명 유지).
        /// 새 메일: 현재 작성 중인 메일에 본문 삽입.
        /// subject가 있으면 제목도 설정.
        /// </summary>
        public void ApplyComposeDraft(string body, string subject = null)
        {
            var inspector = _app.ActiveInspector();
            if (inspector?.CurrentItem is MailItem composing)
            {
                // 현재 작성 중인 메일에 삽입
                if (!string.IsNullOrEmpty(subject))
                    composing.Subject = subject;

                // HTML 본문이 있으면 서명 앞에 삽입, 없으면 덮어쓰기
                if (!string.IsNullOrEmpty(composing.HTMLBody) && composing.HTMLBody.Contains("<body"))
                {
                    // 서명 전에 내용 삽입 (body 태그 바로 뒤)
                    var bodyTag = composing.HTMLBody.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                    var bodyClose = composing.HTMLBody.IndexOf(">", bodyTag) + 1;
                    composing.HTMLBody = composing.HTMLBody.Insert(bodyClose,
                        "<div style='font-family:Calibri,sans-serif;font-size:11pt'>"
                        + body.Replace("\n", "<br>") + "</div><br>");
                }
                else
                {
                    composing.Body = body;
                }
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

                    // Reply는 이미 서명이 포함되어 있으므로 본문만 앞에 추가
                    reply.Body = body + "\n\n" + reply.Body;
                    reply.Display(false);
                }
            }
        }

        public void InsertReplyDraft(string replyBody)
        {
            ApplyComposeDraft(replyBody);
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
