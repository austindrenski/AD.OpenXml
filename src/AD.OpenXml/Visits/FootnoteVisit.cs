using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.OpenXml.Visitors;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <inheritdoc />
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public sealed class FootnoteVisit : IOpenXmlVisit
    {
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] private static readonly IEnumerable<XName> Revisions =
            new XName[]
            {
                W + "ins",
                W + "del",
                W + "rPrChange",
                W + "moveToRangeStart",
                W + "moveToRangeEnd",
                W + "moveTo"
            };

        /// <inheritdoc />
        public IOpenXmlVisitor Result { get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">
        /// The file from which content is copied.
        /// </param>
        /// <param name="footnoteId">
        /// The last footnote number currently in use by the container.
        /// </param>
        /// <param name="revisionId">
        /// The current revision number incremented by one.
        /// </param>
        /// <returns>
        /// The updated document node of the source file.
        /// </returns>
        public FootnoteVisit(IOpenXmlVisitor subject, int footnoteId, int revisionId)
        {
            (XElement document, XElement footnotes) = Execute(subject.Footnotes, subject.Document, footnoteId, revisionId);

            Result =
                new OpenXmlVisitor(
                    subject.ContentTypes,
                    document,
                    subject.DocumentRelations,
                    footnotes,
                    subject.FootnoteRelations,
                    subject.Styles,
                    subject.Numbering,
                    subject.Theme1,
                    subject.Charts);
        }

        [Pure]
        private static (XElement Document, XElement Footnotes) Execute([NotNull] XElement footnotes, [NotNull] XElement document, int footnoteId, int revisionId)
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
                footnotes
                    // Remove editing attributes.
                    .RemoveRsidAttributes()

                    // Remove run properties from the paragraph scope.
                    .RemoveRunPropertiesFromParagraphProperties()

                    // Remove elements that should never exist in-line.
                    .RemoveByAll(W + "bCs")
                    .RemoveByAll(W + "bookmarkEnd")
                    .RemoveByAll(W + "bookmarkStart")
                    .RemoveByAll(W + "color")
                    .RemoveByAll(W + "hideMark")
                    .RemoveByAll(W + "iCs")
                    .RemoveByAll(W + "keepNext")
                    .RemoveByAll(W + "lang")
                    .RemoveByAll(W + "lastRenderedPageBreak")
                    .RemoveByAll(W + "noProof")
                    .RemoveByAll(W + "noWrap")
                    .RemoveByAll(W + "numPr")
                    .RemoveByAll(W + "proofErr")
                    .RemoveByAll(W + "rFonts")
                    .RemoveByAll(W + "spacing")
                    .RemoveByAll(W + "sz")
                    .RemoveByAll(W + "szCs")
                    .RemoveByAll(W + "tblPrEx")
                    .RemoveByAll(W + "commentRangeStart")
                    .RemoveByAll(W + "commentRangeEnd")
                    .RemoveByAll(W + "commentReference")
                    .RemoveByAll(x => (string) x.Attribute(W + "val") == "CommentReference")

                    // Remove elements that should almost never exist.
                    .RemoveByAll(x => x.Name == W + "br" && (string) x.Attribute(W + "type") == "page")
                    .RemoveByAll(x => x.Name == W + "pStyle" && (string) x.Attribute(W + "val") == "BodyTextSSFinal")
                    .RemoveByAll(x => x.Name == W + "pStyle" && (string) x.Attribute(W + "val") == "Default")
                    .RemoveByAll(x => x.Name == W + "jc" && !x.Ancestors(W + "tbl").Any())

                    // Alter bold, italic, and underline elements.
                    .ChangeBoldToStrong()
                    .ChangeItalicToEmphasis()
                    .ChangeUnderlineToTableCaption()
                    .ChangeUnderlineToFigureCaption()
                    .ChangeUnderlineToSourceNote()
                    .ChangeSuperscriptToReference()

                    // Mark insert requests for the production team.
                    .MergeRuns()
                    .HighlightInsertRequests()

                    // Set table styles.
                    .SetTableStyles()

                    // Remove elements used above, but not needed in the output.
                    .RemoveByAll(W + "u")
                    .RemoveByAllIfEmpty(W + "tcPr")
                    .RemoveByAllIfEmpty(W + "rPr")
                    .RemoveByAllIfEmpty(W + "pPr")
                    .RemoveByAllIfEmpty(W + "t")
                    .RemoveByAllIfEmpty(W + "r")
                    .RemoveByAll(x => x.Name == W + "p" && !x.HasElements && x.Parent?.Name != W + "tc")
                    .RemoveBy(x => (int) x.Attribute(W + "id") < 1);

            modifiedFootnotes.Descendants(W + "p")
                             .Attributes()
                             .Remove();

            // There shouldn't be more than one run style.
            foreach (XElement runProperties in modifiedFootnotes.Descendants(W + "rPr").Where(x => x.Elements(W + "rStyle").Count() > 1))
            {
                XElement[] styles =
                    runProperties.Elements(W + "rStyle")
                                 .ToArray();

                styles.Remove();

                IEnumerable<XElement> distinct =
                    styles.Distinct(XNode.EqualityComparer)
                          .Cast<XElement>()
                          .ToArray();

                if (distinct.Any(x => (string) x.Attribute(W + "val") == "FootnoteReference"))
                {
                    distinct = distinct.Where(x => (string) x.Attribute(W + "val") == "FootnoteReference");
                }

                runProperties.AddFirst(distinct);
            }

            (int oldId, int newId)[] footnoteMapping =
                modifiedFootnotes.Elements(W + "footnote")
                                 .Select(x => (int) x.Attribute(W + "id"))
                                 .OrderBy(x => x)
                                 .Select((x, i) => (oldId: x, newId: i + footnoteId))
                                 .ToArray();

            foreach ((int oldId, int newId) in footnoteMapping)
            {
                document.Descendants(W + "footnoteReference").Attributes(W + "id").SingleOrDefault(x => (int) x == oldId)?.SetValue(newId);
                modifiedFootnotes.Descendants(W + "footnote").Attributes(W + "id").SingleOrDefault(x => (int) x == oldId)?.SetValue(newId);
            }

            (int oldId, int newId)[] revisionMapping =
                modifiedFootnotes.Descendants()
                                 .Where(x => Revisions.Contains(x.Name))
                                 .Select(x => (int) x.Attribute(W + "id"))
                                 .OrderBy(x => x)
                                 .Select(x => (oldId: x, newId: x + revisionId))
                                 .ToArray();

            foreach (XName revision in Revisions)
            {
                foreach ((int oldId, int newId) in revisionMapping)
                {
                    modifiedFootnotes = modifiedFootnotes.ChangeXAttributeValues(revision, W + "id", oldId.ToString(), newId.ToString());
                }
            }

            XElement resultFootnotes =
                new XElement(
                    modifiedFootnotes.Name,
                    modifiedFootnotes.Elements()
                                     .OrderBy(x => (int) x.Attribute(W + "id")));

            return (document, resultFootnotes);
        }
    }
}