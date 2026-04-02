using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Models;
using GptOutlookPlugin.Services;

namespace GptOutlookPlugin.Tests.Services
{
    [TestClass]
    public class CodexCliProviderTests
    {
        [TestMethod]
        public void BuildPromptArg_FormatsMessagesCorrectly()
        {
            var provider = new CodexCliProvider(new CodexCliSettings());
            var messages = new List<ChatMessage>
            {
                new ChatMessage(ChatRole.System, "You are helpful."),
                new ChatMessage(ChatRole.User, "Hello")
            };

            var arg = provider.BuildPromptArgument(messages);

            StringAssert.Contains(arg, "You are helpful.");
            StringAssert.Contains(arg, "Hello");
            StringAssert.Contains(arg, "[System]:");
            StringAssert.Contains(arg, "[User]:");
        }

        [TestMethod]
        public void ParseResponse_ExtractsContent()
        {
            var provider = new CodexCliProvider(new CodexCliSettings());
            var raw = "Some preamble\n\nHere is the actual response content.\nWith multiple lines.";

            var result = provider.ParseResponse(raw);

            Assert.IsFalse(string.IsNullOrWhiteSpace(result));
            Assert.AreEqual(raw.Trim(), result);
        }

        [TestMethod]
        public void ParseResponse_HandlesNull()
        {
            var provider = new CodexCliProvider(new CodexCliSettings());

            var result = provider.ParseResponse(null);

            Assert.AreEqual("", result);
        }
    }
}
