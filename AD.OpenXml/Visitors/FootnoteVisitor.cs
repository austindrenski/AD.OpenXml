using System;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public class FootnoteVisitor
    {
        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        public XElement Document { get; }

        /// <summary>
        /// 
        /// </summary>
        public XElement Footnotes { get; }

        /// <summary>
        /// 
        /// </summary>
        public int FootnoteId { get; }
        
        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="content">The document node of the source file containing any modifications made to this point.</param>
        /// <param name="footnoteId">The last footnote number currently in use by the container.</param>
        /// <returns>The updated document node of the source file.</returns>
        public FootnoteVisitor(OpenXmlVisitor subject, XElement content, int footnoteId)
        {
            (Document, Footnotes, FootnoteId) = Execute(subject.Footnotes, content, footnoteId);
        }

        [Pure]
        private static (XElement Document, XElement Footnotes, int FootnoteId) Execute(XElement footnotes, XElement document, int footnoteId)
        {
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            XElement sourceFootnotes =
                footnotes.RemoveRsidAttributes();

            if (sourceFootnotes is null)
            {
                return (document, new XElement(W + "footnotes"), footnoteId);
            }

            sourceFootnotes.Descendants(W + "p")
                           .Attributes()
                           .Remove();

            //sourceFootnotes.Descendants(W + "hyperlink")
            //               .Remove();

            var footnoteMapping =
                sourceFootnotes.Elements(W + "footnote")
                               .Attributes(W + "id")
                               .Select(x => x.Value)
                               .Select(int.Parse)
                               .Where(x => x > 0)
                               .OrderByDescending(x => x)
                               .Select(
                                   x => new
                                   {
                                       oldId = $"{x}",
                                       newId = $"{footnoteId + x}",
                                       newNumericId = footnoteId + x
                                   })
                               .ToArray();

            XElement modifiedContent = document.Clone();

            XElement modifiedFootnotes = sourceFootnotes.Clone();

            foreach (var map in footnoteMapping)
            {
                modifiedContent =
                    modifiedContent.ChangeXAttributeValues(W + "footnoteReference", W + "Id", map.oldId, map.newId);

                modifiedFootnotes =
                    modifiedFootnotes.ChangeXAttributeValues(W + "footnote", W + "id", map.oldId, map.newId);
            }

            int newCurrentId = footnoteMapping.Any() ? footnoteMapping.Max(x => x.newNumericId) : footnoteId;

            return (modifiedContent, modifiedFootnotes, newCurrentId);
        }
    }
}