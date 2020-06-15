using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MsWordTemplateProcessor
{
    public static class Utils
    {
        private static IEnumerable<T> Just<T>(T element)
        {
            yield return element;
        }

        public static Dictionary<string, List<T>> BookmarkValues<T>(
            WordprocessingDocument doc, Func<BookmarkStart, T> func
        )
        {
            var mainDocumentPart = doc.MainDocumentPart;
            var headers = mainDocumentPart.HeaderParts.SelectMany(x => x.Header);
            var footers = mainDocumentPart.FooterParts.SelectMany(x => x.Footer);
            var body = Just(mainDocumentPart.RootElement);
            return headers.Concat(body).Concat(footers)
                .SelectMany(x => x.Descendants<BookmarkStart>())
                .GroupBy(x => x.Name.Value, func)
                .ToDictionary(x => x.Key, x => x.ToList());
        }

        // cross-element bookmarks are not supported
        // bookmarks in table cells are not supported
        internal static IEnumerable<OpenXmlElement> BookmarkContent(BookmarkStart bookmarkStart)
        {
            return bookmarkStart?.ElementsAfter().Where(x => !(x is BookmarkEnd));
        }
    }
}