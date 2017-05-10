using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO.Standard;
using AD.OpenXml.Standard.Elements;
using AD.OpenXml.Standard.Visitors;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Visits
{
    /// <summary>
    /// Marshals content from the 'document.xml' file of a Word document as an idiomatic XML object.
    /// </summary>
    [PublicAPI]
    public sealed class DocumentVisit : IOpenXmlVisit
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
        /// Marshals content from the source document to be added into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="revisionId">
        /// The current revision number incremented by one.
        /// </param>
        /// <returns>The updated document node of the source file.</returns>
        public DocumentVisit(IOpenXmlVisitor subject, int revisionId)
        {
            XElement document = Execute(subject.Document, revisionId);

            Result =
                new OpenXmlVisitor(
                    subject.ContentTypes,
                    document,
                    subject.DocumentRelations,
                    subject.Footnotes,
                    subject.FootnoteRelations,
                    subject.Styles,
                    subject.Numbering,
                    subject.Charts);
        }

        [Pure]
        [NotNull]
        private static XElement Execute([NotNull] XElement document, int revisionId)
        {
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            XElement source =
                document

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
                    .RemoveByAll(W + "shd")
                    .RemoveByAll(W + "spacing")
                    .RemoveByAll(W + "sz")
                    .RemoveByAll(W + "szCs")
                    .RemoveByAll(W + "tblPrEx")
                    .RemoveByAll(W + "commentRangeStart")
                    .RemoveByAll(W + "commentRangeEnd")
                    .RemoveByAll(W + "commentReference")
                    .RemoveByAll(x => (string) x.Attribute(W + "val") == "CommentReference")

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

                    // Remove for this stage
                    .RemoveByAll(W + "footerReference")
                    .RemoveByAll(W + "headerReference")

                    // Add soft breaks to headings
                    .AddLineBreakToHeadings()

                    // Tidy up the XML for review.
                    .MergeRuns();

            // There shouldn't be more than one paragraph style.
            foreach (XElement paragraphProperties in source.Descendants(W + "pPr").Where(x => x.Elements(W + "pStyle").Count() > 1))
            {
                IEnumerable<XElement> styles = paragraphProperties.Elements(W + "pStyle").ToArray();
                styles.Remove();
                paragraphProperties.AddFirst(styles.Distinct());
            }

            // There shouldn't be more than one run style.
            foreach (XElement runProperties in source.Descendants(W + "rPr").Where(x => x.Elements(W + "rStyle").Count() > 1))
            {
                IEnumerable<XElement> styles = runProperties.Elements(W + "rStyle").ToArray();
                styles.Remove();
                IEnumerable<XElement> distinct = styles.Distinct(XNode.EqualityComparer).Cast<XElement>().ToArray();
                if (distinct.Any(x => x.Attribute(W + "val")?.Value.Equals("FootnoteReference") ?? false))
                {
                    distinct = distinct.Where(x => x.Attribute(W + "val")?.Value.Equals("FootnoteReference") ?? false);
                }
                runProperties.AddFirst(distinct);
            }

            var revisionMapping =
                source.Descendants()
                      .Where(x => Revisions.Contains(x.Name))
                      .Attributes(W + "id")
                      .OrderBy(x => x.Value.ParseInt())
                      .Select(
                          x => new
                          {
                              oldId = x,
                              newId = new XAttribute(W + "id", $"{revisionId + x.Value.ParseInt()}")
                          })
                      .OrderByDescending(x => x.oldId.Value.ParseInt())
                      .ToArray();

            foreach (XName revision in Revisions)
            {
                foreach (var map in revisionMapping)
                {
                    source = source.ChangeXAttributeValues(revision, W + "id", (string)map.oldId, (string)map.newId);
                }
            }

            source.Descendants(W + "sectPr").Attributes().Remove();

            source.Descendants(W + "p").Attributes().Remove();

            source.Descendants(W + "tr").Attributes().Remove();

            if (source.Element(W + "body")?.Elements().First().Name == W + "sectPr")
            {
                source.Element(W + "body")?.Elements().First().Remove();
            }

            if (source.Element(W + "body")?.Elements().Last().Name == W + "sectPr")
            {
                XElement sectionProperties = source.Element(W + "body")?.Elements().Last();

                XElement previous = sectionProperties.Previous();

                sectionProperties?.Remove();

                if (!previous?.Elements(W + "pPr").Any() ?? false)
                {
                    previous?.AddFirst(new XElement(W + "pPr", sectionProperties));
                }
                previous?.Element(W + "pPr")?.Add(sectionProperties);
            }

            return source;
        }
    }
}