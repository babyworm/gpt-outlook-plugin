using System;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Core
{
    public static class PromptTemplates
    {
        public static string GetSystemPrompt(FeatureMode mode, EmailContext ctx,
            string targetLanguage = "Korean", string userLanguage = "Korean",
            string userName = "", string userEmail = "",
            string tone = "Professional and polite",
            string reviewSensitivity = "Medium")
        {
            string emailSection = $"\n\n---\nEmail Subject: {ctx.Subject}\nRecipients: {ctx.Recipients}\nEmail Body:\n{ctx.Body}\n---";

            string respondIn = $"Always respond in {userLanguage}.";
            string toneInstruction = $"Use a {tone} tone.";

            switch (mode)
            {
                case FeatureMode.Rewrite:
                    return "You are an expert email rewriter. "
                         + "Rewrite the following email in two alternative versions:\n\n"
                         + "## Version 1: Natural\n"
                         + "A natural, fluent version that sounds like a native speaker wrote it. "
                         + "Fix all grammar issues and improve readability.\n\n"
                         + "## Version 2: Business Professional\n"
                         + "A polished, business-appropriate version with proper formality. "
                         + toneInstruction + "\n\n"
                         + "Present each version with the header above. "
                         + "Keep the same meaning and intent as the original. "
                         + respondIn
                         + emailSection;

                case FeatureMode.Proofread:
                    var sensitivityGuide = "";
                    switch (reviewSensitivity)
                    {
                        case "Low":
                            sensitivityGuide = "Only flag major issues: grammar errors, factual mistakes, or confusing sentences. Ignore minor style preferences. ";
                            break;
                        case "High":
                            sensitivityGuide = "Flag all issues including minor style improvements, word choice, tone adjustments, and formatting. ";
                            break;
                        default: // Medium
                            sensitivityGuide = "Flag grammar errors, clarity issues, and awkward phrasing. Skip trivial style preferences. ";
                            break;
                    }
                    return "You are an expert email proofreader. "
                         + "Proofread the following email for grammar, tone, clarity, and professionalism. "
                         + $"The desired tone is: {tone}. "
                         + sensitivityGuide
                         + "You MUST output EACH correction using this EXACT block format. "
                         + "Do NOT use any other format. Do NOT add extra commentary outside the blocks.\n\n"
                         + "[CORRECTION]\n"
                         + "ORIGINAL: <copy the exact original text that needs fixing>\n"
                         + "CORRECTED: <the corrected version>\n"
                         + "REASON: <brief explanation>\n"
                         + "[/CORRECTION]\n\n"
                         + "Output one [CORRECTION]...[/CORRECTION] block per issue found. "
                         + "If the email has no issues, respond with exactly: [NO_ISSUES] The email looks good. "
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
                         + "First, detect the language of the email below. "
                         + $"If the email is in {targetLanguage} (or very close), translate it to English. "
                         + $"If the email is in any other language, translate it to {targetLanguage}. "
                         + "Maintain the original tone, formality, and formatting. "
                         + "Output only the translated text without additional explanation. "
                         + "Do NOT include the original text or language detection notes."
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
