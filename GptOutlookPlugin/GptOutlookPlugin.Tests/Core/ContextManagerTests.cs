using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Core;
using GptOutlookPlugin.Models;

namespace GptOutlookPlugin.Tests.Core
{
    [TestClass]
    public class ContextManagerTests
    {
        [TestMethod]
        public void GetOrCreateSession_CreatesNewSession()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var session = mgr.GetOrCreateSession("email-001", FeatureMode.Review);

            Assert.IsNotNull(session);
            Assert.AreEqual("email-001", session.SessionKey);
            Assert.AreEqual(FeatureMode.Review, session.CurrentMode);
        }

        [TestMethod]
        public void GetOrCreateSession_ReturnsSameSession()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var s1 = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            var s2 = mgr.GetOrCreateSession("email-001", FeatureMode.Review);

            Assert.AreSame(s1, s2);
        }

        [TestMethod]
        public void GetOrCreateSession_DifferentEmailsGetDifferentSessions()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var s1 = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            var s2 = mgr.GetOrCreateSession("email-002", FeatureMode.Translate);

            Assert.AreNotSame(s1, s2);
        }

        [TestMethod]
        public void GetOrCreateSession_SameEmailDifferentMode_SwitchesMode()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var s1 = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            s1.AddMessage(ChatRole.User, "Review this");

            var s2 = mgr.GetOrCreateSession("email-001", FeatureMode.Translate);

            Assert.AreSame(s1, s2);
            Assert.AreEqual(FeatureMode.Translate, s2.CurrentMode);
            Assert.AreEqual(1, s2.Messages.Count);
        }

        [TestMethod]
        public void ClearSession_RemovesSession()
        {
            var mgr = new ContextManager(maxHistory: 10);
            mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            mgr.ClearSession("email-001");

            var session = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            Assert.AreEqual(0, session.Messages.Count);
        }

        [TestMethod]
        public void BuildMessages_IncludesSystemPromptAndHistory()
        {
            var mgr = new ContextManager(maxHistory: 10);
            var session = mgr.GetOrCreateSession("email-001", FeatureMode.Review);
            session.EmailContext = new EmailContext { Subject = "Test", Body = "Hello" };
            session.AddMessage(ChatRole.User, "Review this");
            session.AddMessage(ChatRole.Assistant, "Looks good");

            var messages = mgr.BuildMessages(session);

            Assert.AreEqual(ChatRole.System, messages[0].Role);
            Assert.AreEqual(3, messages.Count);
        }
    }
}
