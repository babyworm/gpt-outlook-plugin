using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Core;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Tests.Core
{
    [TestClass]
    public class PromptTemplatesTests
    {
        [TestMethod]
        public void GetSystemPrompt_Review_ContainsEmailBody()
        {
            var ctx = new EmailContext { Subject = "Test", Body = "Hello world" };
            var prompt = PromptTemplates.GetSystemPrompt(FeatureMode.Review, ctx);

            StringAssert.Contains(prompt, "Hello world");
            StringAssert.Contains(prompt, "proofread");
        }

        [TestMethod]
        public void GetSystemPrompt_Translate_ContainsTargetLanguage()
        {
            var ctx = new EmailContext { Body = "Hello" };
            var prompt = PromptTemplates.GetSystemPrompt(FeatureMode.Translate, ctx, targetLanguage: "ko");

            StringAssert.Contains(prompt, "ko");
        }

        [TestMethod]
        public void GetSystemPrompt_Compose_ContainsRecipient()
        {
            var ctx = new EmailContext
            {
                Body = "Original email",
                Recipients = "boss@company.com"
            };
            var prompt = PromptTemplates.GetSystemPrompt(FeatureMode.Compose, ctx);

            StringAssert.Contains(prompt, "boss@company.com");
        }

        [TestMethod]
        public void GetSystemPrompt_AllModes_ReturnNonEmpty()
        {
            var ctx = new EmailContext { Body = "test" };

            foreach (FeatureMode mode in System.Enum.GetValues(typeof(FeatureMode)))
            {
                var prompt = PromptTemplates.GetSystemPrompt(mode, ctx);
                Assert.IsFalse(string.IsNullOrWhiteSpace(prompt),
                    $"Prompt for {mode} should not be empty");
            }
        }
    }
}
