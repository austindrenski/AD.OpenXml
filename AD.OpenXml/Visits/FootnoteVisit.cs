using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.OpenXml.Visitors;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public sealed class FootnoteVisit : IOpenXmlVisit
    {
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;
        
        [NotNull]
        private static readonly IEnumerable<XName> Revisions =
            new XName[]
            {
                W + "ins",
                W + "del",
                W + "rPrChange",
                W + "moveToRangeStart",
                W + "moveToRangeEnd",
                W + "moveTo"
            };

        /// <summary>
        /// 
        /// </summary>
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
            (var document, var footnotes) = Execute(subject.Footnotes, subject.Document, footnoteId, revisionId);

             Result =
                new OpenXmlVisitor(
                    subject.ContentTypes,
                    document,
                    subject.DocumentRelations,
                    footnotes,
                    subject.FootnoteRelations,
                    subject.Styles,
                    subject.Numbering,
                    subject.Charts);
        }

        [Pure]
        private static (XElement Document, XElement Footnotes) Execute(XElement footnotes, XElement document, int footnoteId, int revisionId)
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
                    .RemoveByAll(x => (string)x.Attribute(W + "val") == "CommentReference")
                    .RemoveByAll(W + "author")

                    // Remove elements that should almost never exist.
                    .RemoveByAll(x => x.Name.Equals(W + "br") && (x.Attribute(W + "type")?.Value.Equals("page", StringComparison.OrdinalIgnoreCase) ?? false))
                    .RemoveByAll(x => x.Name.Equals(W + "pStyle") && (x.Attribute(W + "val")?.Value.Equals("BodyTextSSFinal", StringComparison.OrdinalIgnoreCase) ?? false))
                    .RemoveByAll(x => x.Name.Equals(W + "pStyle") && (x.Attribute(W + "val")?.Value.Equals("Default", StringComparison.OrdinalIgnoreCase) ?? false))
                    .RemoveByAll(x => x.Name.Equals(W + "jc") && !x.Ancestors(W + "tbl").Any())

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
                    .RemoveByAll(x => x.Name.Equals(W + "p") && !x.HasElements && (!x.Parent?.Name.Equals(W + "tc") ?? false))

                    .RemoveBy(x => int.Parse(x.Attribute(W + "id")?.Value ?? "0") < 1);

            modifiedFootnotes.Descendants(W + "p")
                             .Attributes()
                             .Remove();

            IEnumerable<(string oldId, string newId)> footnoteMapping =
                modifiedFootnotes.Elements(W + "footnote")
                                 .Select(
                                    x => x.Attribute(W + "id"))
                                 .OrderBy(
                                    x => x?.Value.ParseInt())
                                 .Select(
                                     (x, i) => (oldId: x.Value, newId: $"{footnoteId + i}"))
                                 .OrderByDescending(x => x.oldId.ParseInt())
                                 .ToArray();

            foreach ((string oldId, string newId) map in footnoteMapping)
            {
                document.Descendants(W + "footnoteReference").Attributes(W + "id").Single(x => x.Value == map.oldId).SetValue(map.newId);
                modifiedFootnotes.Descendants(W + "footnote").Attributes(W + "id").Single(x => x.Value == map.oldId).SetValue(map.newId);
            }

            var revisionMapping =
                modifiedFootnotes.Descendants()
                                 .Where(x => Revisions.Contains(x.Name))
                                 .OrderBy(x => x.Attribute(W + "id")?.Value.ParseInt())
                                 .Select(
                                     x => new
                                     {
                                         type = x.Name,
                                         oldId = x.Attribute(W + "id"),
                                         newId = new XAttribute(W + "id", $"{revisionId + x.Attribute(W + "id")?.Value.ParseInt()}")
                                     })
                                 .OrderByDescending(x => x.oldId.Value.ParseInt())
                                 .ToArray();

            revisionMapping.ToList()
                           .ForEach(x => Console.WriteLine($"type: {x.type}, old: {x.oldId}, new: {x.newId}"));

            foreach (XName revision in Revisions)
            {
                foreach (var map in revisionMapping)
                {
                    modifiedFootnotes = modifiedFootnotes.ChangeXAttributeValues(revision, W + "id", (string) map.oldId, (string) map.newId);
                }
            }

            XElement resultFootnotes =
                new XElement(
                    modifiedFootnotes.Name,
                    modifiedFootnotes.Elements()
                                     .OrderBy(x => int.Parse(x.Attribute(W + "id")?.Value ?? "0")));

            return (document, resultFootnotes);
        }
    }
}