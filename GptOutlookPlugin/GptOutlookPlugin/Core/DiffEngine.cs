using System.Collections.Generic;
using DiffPlex;
using DiffPlex.DiffBuilder;
using DiffPlex.DiffBuilder.Model;

namespace GptOutlookPlugin.Core
{
    public enum DiffType
    {
        Unchanged,
        Inserted,
        Deleted,
        Modified
    }

    public class DiffLine
    {
        public DiffType Type { get; set; }
        public string Text { get; set; }
    }

    public static class DiffEngine
    {
        public static List<DiffLine> ComputeSentenceDiff(string original, string modified)
        {
            var differ = new Differ();
            var builder = new InlineDiffBuilder(differ);
            var diff = builder.BuildDiffModel(original, modified, ignoreWhitespace: false);

            var result = new List<DiffLine>();

            foreach (var line in diff.Lines)
            {
                switch (line.Type)
                {
                    case ChangeType.Unchanged:
                        result.Add(new DiffLine { Type = DiffType.Unchanged, Text = line.Text });
                        break;
                    case ChangeType.Deleted:
                        result.Add(new DiffLine { Type = DiffType.Deleted, Text = line.Text });
                        break;
                    case ChangeType.Inserted:
                        result.Add(new DiffLine { Type = DiffType.Inserted, Text = line.Text });
                        break;
                    case ChangeType.Modified:
                        result.Add(new DiffLine { Type = DiffType.Modified, Text = line.Text });
                        break;
                    case ChangeType.Imaginary:
                        break;
                }
            }

            return result;
        }
    }
}
