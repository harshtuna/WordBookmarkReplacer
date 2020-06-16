using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace MsWordTemplateProcessor.Cli.Test
{
    [TestFixture]
    public class CliTests
    {
        [Test]
        public void TestEmpty()
        {
            RunTest(new string[] { }, new Dictionary<string, List<string>>
            {
                ["hdr_bkm"] = new List<string> {"", ""},
                ["bdy_bkm"] = new List<string> {"Old text", "Old text"},
                ["ftr_bkm"] = new List<string> {"Old footer"},
            });
        }

        [Test]
        public void TestOne()
        {
            RunTest(new[] {"-b", "bdy_bkm", "-v", "test value"}, new Dictionary<string, List<string>>
            {
                ["hdr_bkm"] = new List<string> {"", ""},
                ["bdy_bkm"] = new List<string> {"test value", "test value"},
                ["ftr_bkm"] = new List<string> {"Old footer"},
            });
        }

        [Test]
        public void TestMany()
        {
            RunTest(new[] {"-b", "bdy_bkm", "-v", "test value", "-b", "ftr_bkm", "-v", "test"}, new Dictionary<string, List<string>>
            {
                ["hdr_bkm"] = new List<string> {"", ""},
                ["bdy_bkm"] = new List<string> {"test value", "test value"},
                ["ftr_bkm"] = new List<string> {"test footer"},
            });
        }

        private static void RunTest(string[] values, Dictionary<string, List<string>> expected)
        {
            var templatePath = Path.Combine(TestContext.CurrentContext.TestDirectory, @"Resources\test.docx");
            var file2Path = Path.GetTempFileName() + ".docx";
            Program.Main(new[] {"-t", templatePath, "-o", file2Path}.Concat(values).ToArray());
            using var output = WordprocessingDocument.Open(file2Path, true);
            Assert.AreEqual(
                expected,
                Utils.BookmarkValues(output, bs => bs.Parent.InnerText)
            );
        }
    }
}