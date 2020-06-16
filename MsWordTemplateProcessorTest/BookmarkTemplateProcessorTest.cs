using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace MsWordTemplateProcessor.Test
{
    public class BookmarkReplacerTest
    {
        [Test]
        public void TestInsertTextIntoBookmarkInMemory()
        {
            var templatePath = Path.Combine(TestContext.CurrentContext.TestDirectory, @"Resources\test.docx");

            Console.WriteLine($"templatePath: {templatePath}");

            using var memoryStream = new MemoryStream();
            var templateProcessor = new BookmarkTemplateProcessor(templatePath, memoryStream);

            Assert.AreEqual(
                new Dictionary<string, List<string>>
                {
                    ["hdr_bkm"] = new List<string> {"", ""},
                    ["bdy_bkm"] = new List<string> {"Old text", "Old text"},
                    ["ftr_bkm"] = new List<string> {"Old footer"},
                },
                Utils.BookmarkValues(templateProcessor.Document, bs => bs.Parent.InnerText)
            );

            templateProcessor.ApplyValue("hdr_bkm", "header test");
            templateProcessor.ApplyValues(new Dictionary<string, string> {["bdy_bkm"] = "body test"});

            Assert.AreEqual(
                new Dictionary<string, List<string>>
                {
                    ["hdr_bkm"] = new List<string> {"header test", "header test"},
                    ["bdy_bkm"] = new List<string> {"body test", "body test"},
                    ["ftr_bkm"] = new List<string> {"Old footer"},
                },
                Utils.BookmarkValues(templateProcessor.Document, bs => bs.Parent.InnerText)
            );
        }

        [Test]
        public void TestInsertTextIntoBookmarkInFile()
        {
            var templatePath = Path.Combine(TestContext.CurrentContext.TestDirectory, @"Resources\test.docx");

            Console.WriteLine($"templatePath: {templatePath}");

            var file2Path = Path.GetTempFileName() + ".docx";
            Console.WriteLine($"temp file: {file2Path}");
            var file2Processor1 = new BookmarkTemplateProcessor(templatePath, file2Path);
            file2Processor1.ApplyValues(new Dictionary<string, string>
            {
                ["hdr_bkm"] = "header test", ["bdy_bkm"] = "body test", ["ftr_bkm"] = "new",
            });
            Assert.AreEqual(
                new Dictionary<string, List<string>>
                {
                    ["hdr_bkm"] = new List<string> {"header test", "header test"},
                    ["bdy_bkm"] = new List<string> {"body test", "body test"},
                    ["ftr_bkm"] = new List<string> {"new footer"},
                },
                Utils.BookmarkValues(file2Processor1.Document, bs => bs.Parent.InnerText)
            );
            file2Processor1.Document.Save();
            file2Processor1.Document.Close();
            
            var file2Processor2 = new BookmarkTemplateProcessor(file2Path);
            file2Processor2.ApplyValue("hdr_bkm", "test in place");
            file2Processor2.Document.Save();
            file2Processor2.Document.Close();
            using var document2 = WordprocessingDocument.Open(file2Path, true);
            Assert.AreEqual(
                new Dictionary<string, List<string>>
                {
                    ["hdr_bkm"] = new List<string> {"test in place", "test in place"},
                    ["bdy_bkm"] = new List<string> {"body test", "body test"},
                    ["ftr_bkm"] = new List<string> {"new footer"},
                },
                Utils.BookmarkValues(document2, bs => bs.Parent.InnerText)
            );
        }
    }
}