using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public class OpenXmlFootnoteVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// Active version of 'word/document.xml'.
        /// </summary>
        public override XElement Document { get; }
        
        /// <summary>
        /// Active version of 'word/footnotes.xml'.
        /// </summary>
        public override XElement Footnotes { get; }
        
        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">
        /// The file from which content is copied.
        /// </param>
        /// <param name="footnoteId">
        /// The last footnote number currently in use by the container.
        /// </param>
        /// <returns>
        /// The updated document node of the source file.
        /// </returns>
        public OpenXmlFootnoteVisitor(OpenXmlVisitor subject, int footnoteId) : base(subject)
        {
            (Document, Footnotes) = Execute(subject.Footnotes, subject.Document, footnoteId);
        }

        [Pure]
        private static (XElement Document, XElement Footnotes) Execute(XElement footnotes, XElement document, int footnoteId)
        {
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            XElement modifiedFootnotes =
                footnotes.RemoveRsidAttributes()
                         .RemoveByAll(W + "bookmarkStart")
                         .RemoveByAll(W + "bookmarkEnd")
                         .RemoveBy(x => int.Parse(x.Attribute(W + "id")?.Value ?? "0") < 1);

            modifiedFootnotes.Descendants(W + "p")
                             .Attributes()
                             .Remove();

            IEnumerable<(string oldId, string newId)> footnoteMapping =
                modifiedFootnotes.Elements(W + "footnote")
                                 .Select(x => x.Attribute(W + "id"))
                                 .OrderBy(x => x?.Value.ParseInt())
                                 .Select(
                                     (x, i) => (oldId: x.Value, newId: $"{footnoteId + i}"))
                                 .OrderByDescending(x => x.oldId.ParseInt())
                                 .ToArray();

            foreach ((string oldId, string newId) map in footnoteMapping)
            {
                document =
                    document.ChangeXAttributeValues(W + "footnoteReference", W + "id", map.oldId, map.newId);

                modifiedFootnotes =
                    modifiedFootnotes.ChangeXAttributeValues(W + "footnote", W + "id", map.oldId, map.newId);
            }

            XElement resultFootnotes =
                new XElement(
                    modifiedFootnotes.Name,
                    modifiedFootnotes.Attributes(),
                    modifiedFootnotes.Elements()
                                     .OrderBy(x => int.Parse(x.Attribute(W + "id")?.Value ?? "0")));

            return (document, resultFootnotes);
        }
    }
}