using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Tests.Models
{
    [TestClass]
    public class ConversationSessionTests
    {
        [TestMethod]
        public void NewSession_HasEmptyHistory()
        {
            var session = new ConversationSession("test-id", FeatureMode.Review);
            Assert.AreEqual(0, session.Messages.Count);
            Assert.AreEqual(FeatureMode.Review, session.CurrentMode);
        }

        [TestMethod]
        public void AddMessage_AppendsToHistory()
        {
            var session = new ConversationSession("test-id", FeatureMode.Review);
            session.AddMessage(ChatRole.User, "Review this email");
            session.AddMessage(ChatRole.Assistant, "Here is the review...");

            Assert.AreEqual(2, session.Messages.Count);
            Assert.AreEqual(ChatRole.User, session.Messages[0].Role);
            Assert.AreEqual(ChatRole.Assistant, session.Messages[1].Role);
        }

        [TestMethod]
        public void TrimHistory_KeepsMaxMessages()
        {
            var session = new ConversationSession("test-id", FeatureMode.Review, maxHistory: 3);

            for (int i = 0; i < 5; i++)
                session.AddMessage(ChatRole.User, $"Message {i}");

            Assert.AreEqual(3, session.Messages.Count);
            Assert.AreEqual("Message 2", session.Messages[0].Content);
            Assert.AreEqual("Message 4", session.Messages[2].Content);
        }

        [TestMethod]
        public void SetEmailContext_UpdatesContext()
        {
            var session = new ConversationSession("test-id", FeatureMode.Translate);
            var ctx = new EmailContext
            {
                Subject = "Meeting Tomorrow",
                Body = "Let's meet at 3pm.",
                Recipients = "bob@example.com"
            };

            session.EmailContext = ctx;

            Assert.AreEqual("Meeting Tomorrow", session.EmailContext.Subject);
        }

        [TestMethod]
        public void SwitchMode_ChangesCurrentMode()
        {
            var session = new ConversationSession("test-id", FeatureMode.Review);
            session.SwitchMode(FeatureMode.Translate);

            Assert.AreEqual(FeatureMode.Translate, session.CurrentMode);
        }
    }
}
