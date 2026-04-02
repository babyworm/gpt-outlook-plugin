using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using GptOutlookPlugin.Core;

namespace GptOutlookPlugin.Tests.Core
{
    [TestClass]
    public class DiffEngineTests
    {
        [TestMethod]
        public void ComputeDiff_IdenticalTexts_AllUnchanged()
        {
            var result = DiffEngine.ComputeSentenceDiff("Hello world.", "Hello world.");

            Assert.AreEqual(1, result.Count);
            Assert.AreEqual(DiffType.Unchanged, result[0].Type);
        }

        [TestMethod]
        public void ComputeDiff_ModifiedSentence_ShowsDeleteAndInsert()
        {
            var original = "I want to talk about the project.";
            var modified = "I would like to discuss the project.";

            var result = DiffEngine.ComputeSentenceDiff(original, modified);

            Assert.IsTrue(result.Any(r => r.Type == DiffType.Deleted));
            Assert.IsTrue(result.Any(r => r.Type == DiffType.Inserted));
        }

        [TestMethod]
        public void ComputeDiff_AddedSentence_ShowsInsert()
        {
            var original = "First sentence.";
            var modified = "First sentence.\nSecond sentence added.";

            var result = DiffEngine.ComputeSentenceDiff(original, modified);

            Assert.IsTrue(result.Any(r => r.Type == DiffType.Inserted));
            Assert.IsTrue(result.Any(r => r.Type == DiffType.Unchanged && r.Text.Contains("First")));
        }

        [TestMethod]
        public void ComputeDiff_RemovedSentence_ShowsDelete()
        {
            var original = "Keep this.\nRemove this.";
            var modified = "Keep this.";

            var result = DiffEngine.ComputeSentenceDiff(original, modified);

            Assert.IsTrue(result.Any(r => r.Type == DiffType.Deleted && r.Text.Contains("Remove")));
        }

        [TestMethod]
        public void ComputeDiff_EmptyInputs_ReturnsEmpty()
        {
            var result = DiffEngine.ComputeSentenceDiff("", "");
            Assert.IsNotNull(result);
        }
    }
}
