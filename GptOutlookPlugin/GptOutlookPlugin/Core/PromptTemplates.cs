using System;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public static class PromptTemplates
    {
        public static string GetSystemPrompt(FeatureMode mode, EmailContext ctx,
            string targetLanguage = "Korean", string userLanguage = "Korean",
            string userName = "", string userEmail = "",
            string tone = "Professional and polite")
        {
            string emailSection = $"\n\n---\nEmail Subject: {ctx.Subject}\nRecipients: {ctx.Recipients}\nEmail Body:\n{ctx.Body}\n---";

            // 모든 모드에서 사용자 언어로 응답하도록 지시
            string respondIn = $"Always respond in {userLanguage}.";
            string toneInstruction = $"Use a {tone} tone.";

            switch (mode)
            {
                case FeatureMode.Review:
                    return "You are an expert email proofreader. "
                         + "Review the following email for grammar, tone, clarity, and professionalism. "
                         + $"The desired tone is: {tone}. "
                         + "Present your corrections in a structured before/after format. "
                         + "For each change, briefly explain why. "
                         + respondIn
                         + emailSection;

                case FeatureMode.Compose:
                    var senderInfo = !string.IsNullOrEmpty(userName)
                        ? $"The sender's name is {userName} ({userEmail}). "
                        : "";
                    return "You are an expert business email writer. "
                         + senderInfo
                         + $"Draft an appropriate reply to the following email. Consider the recipient ({ctx.Recipients}) and context. "
                         + "Match the formality level of the original email. "
                         + toneInstruction + " "
                         + "IMPORTANT: Do NOT include greetings (Dear...), closings (Best regards...), "
                         + "or signatures — Outlook will add those automatically. "
                         + "Output the email as two clearly separated sections:\n"
                         + "Subject: <subject line>\n\n<email body>\n"
                         + "Write the reply in the same language as the original email. "
                         + $"Provide your explanations in {userLanguage}."
                         + emailSection;

                case FeatureMode.AutoReply:
                    var replyerInfo = !string.IsNullOrEmpty(userName)
                        ? $"The replyer's name is {userName} ({userEmail}). "
                        : "";
                    return "You are an expert business email writer. "
                         + replyerInfo
                         + $"Draft a concise, appropriate reply to the following email. Consider the recipient ({ctx.Recipients}) and context. "
                         + toneInstruction + " "
                         + "IMPORTANT: Do NOT include greetings (Dear...), closings (Best regards...), "
                         + "or signatures — Outlook will add those automatically. "
                         + "Output the email as two clearly separated sections:\n"
                         + "Subject: <subject line>\n\n<email body>\n"
                         + respondIn
                         + emailSection;

                case FeatureMode.Translate:
                    return "You are a professional translator specializing in business communication. "
                         + $"Translate the following email to {targetLanguage}. "
                         + "Maintain the original tone, formality, and formatting. "
                         + "Output only the translated text without additional explanation."
                         + emailSection;

                case FeatureMode.Summarize:
                    return "You are an expert email analyst. "
                         + "Summarize the following email concisely. Include: "
                         + "1) Key points, 2) Action items or requests, 3) Deadlines if any. "
                         + respondIn
                         + emailSection;

                default:
                    throw new ArgumentOutOfRangeException(nameof(mode));
            }
        }
    }
}
