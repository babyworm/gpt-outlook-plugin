using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using GptOutlookPlugin.Models;
using GptOutlookPlugin.Services;

namespace GptOutlookPlugin.Tests.Services
{
    [TestClass]
    public class AiServiceManagerTests
    {
        [TestMethod]
        public async Task SendAsync_UsesPrimaryWhenAvailable()
        {
            var primary = new Mock<IAiProvider>();
            primary.Setup(p => p.IsAvailable()).Returns(true);
            primary.Setup(p => p.Name).Returns("Primary");
            primary.Setup(p => p.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()))
                   .ReturnsAsync("primary response");

            var fallback = new Mock<IAiProvider>();

            var manager = new AiServiceManager(primary.Object, fallback.Object);
            var messages = new List<ChatMessage> { new ChatMessage(ChatRole.User, "test") };

            var result = await manager.SendAsync(messages, CancellationToken.None);

            Assert.AreEqual("primary response", result);
            fallback.Verify(f => f.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()), Times.Never);
        }

        [TestMethod]
        public async Task SendAsync_FallsBackWhenPrimaryFails()
        {
            var primary = new Mock<IAiProvider>();
            primary.Setup(p => p.IsAvailable()).Returns(true);
            primary.Setup(p => p.Name).Returns("Primary");
            primary.Setup(p => p.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()))
                   .ThrowsAsync(new System.Exception("primary failed"));

            var fallback = new Mock<IAiProvider>();
            fallback.Setup(p => p.IsAvailable()).Returns(true);
            fallback.Setup(p => p.Name).Returns("Fallback");
            fallback.Setup(p => p.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync("fallback response");

            var manager = new AiServiceManager(primary.Object, fallback.Object);
            var messages = new List<ChatMessage> { new ChatMessage(ChatRole.User, "test") };

            var result = await manager.SendAsync(messages, CancellationToken.None);

            Assert.AreEqual("fallback response", result);
        }

        [TestMethod]
        public async Task SendAsync_SkipsPrimaryWhenUnavailable()
        {
            var primary = new Mock<IAiProvider>();
            primary.Setup(p => p.IsAvailable()).Returns(false);
            primary.Setup(p => p.Name).Returns("Primary");

            var fallback = new Mock<IAiProvider>();
            fallback.Setup(p => p.IsAvailable()).Returns(true);
            fallback.Setup(p => p.Name).Returns("Fallback");
            fallback.Setup(p => p.SendAsync(It.IsAny<List<ChatMessage>>(), It.IsAny<CancellationToken>()))
                    .ReturnsAsync("fallback response");

            var manager = new AiServiceManager(primary.Object, fallback.Object);
            var messages = new List<ChatMessage> { new ChatMessage(ChatRole.User, "test") };

            var result = await manager.SendAsync(messages, CancellationToken.None);

            Assert.AreEqual("fallback response", result);
        }
    }
}
