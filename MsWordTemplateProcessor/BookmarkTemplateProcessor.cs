using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MsWordTemplateProcessor
{
    public class BookmarkTemplateProcessor
    {
        public WordprocessingDocument Document { get; }
        private readonly Dictionary<string, List<BookmarkStart>> _bookmarks;

        // must receive writeable document (properly writeable - see OpenXmlPackage.CanSave) 
        private BookmarkTemplateProcessor(WordprocessingDocument document)
        {
            Document = document;
            _bookmarks = Utils.BookmarkValues(document, x => x);
        }

        public BookmarkTemplateProcessor(string filePath) : this(OpenDocument(filePath))
        {
        }

        public BookmarkTemplateProcessor(string templatePath, string filePath) : this(
            NewDocumentUsingTemplate(templatePath, filePath)
        )
        {
        }

        public BookmarkTemplateProcessor(string templatePath, Stream stream) : this(
            NewDocumentUsingTemplate(templatePath, stream)
        )
        {
        }

        ~BookmarkTemplateProcessor()
        {
            try
            {
                Document.Close();
            }
            catch (ObjectDisposedException)
            {
            }
        }

        // 
        public void ApplyValues(IDictionary<string, string> bookmarkToNewText)
        {
            foreach (var (bookmarkName, newText) in bookmarkToNewText.Select(x => (x.Key, x.Value)))
                InsertTextIntoBookmark(bookmarkName, newText);
        }

        public void ApplyValue(string bookmarkName, string newText)
        {
            InsertTextIntoBookmark(bookmarkName, newText);
        }

        private void InsertTextIntoBookmark(string bookmarkName, string newText)
        {
            foreach (var bookmarkStart in _bookmarks[bookmarkName])
            {
                EraseBookmarkContent(bookmarkStart);
                bookmarkStart?.Parent.InsertAfter(new Run(new Text(newText)), bookmarkStart);
            }
        }

        private static WordprocessingDocument OpenDocument(string filePath)
        {
            CheckCanSave();
            return WordprocessingDocument.Open(filePath, true);
        }

        private static WordprocessingDocument NewDocumentUsingTemplate(string templatePath, string filePath)
        {
            CheckCanSave();
            File.Copy(templatePath, filePath, true);
            return WordprocessingDocument.Open(filePath, true);
        }

        private static WordprocessingDocument NewDocumentUsingTemplate(string templatePath, Stream stream)
        {
            // fixme - need to copy each part
            return WordprocessingDocument.Open(templatePath, false)
                .Clone(stream) as WordprocessingDocument;
        }

        private static void CheckCanSave()
        {
            if (!OpenXmlPackage.CanSave)
                throw new PlatformNotSupportedException(
                    "OpenXmlPackage cannot be saved on " + Environment.OSVersion.Platform
                );
        }

        private static void EraseBookmarkContent(BookmarkStart bookmarkStart)
        {
            foreach (var element in Utils.BookmarkContent(bookmarkStart)) element.Remove();
        }
    }
}