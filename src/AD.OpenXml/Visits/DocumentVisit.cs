using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.OpenXml.Structures;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    /// Marshals content from the 'document.xml' file of a Word document as an idiomatic XML object.
    /// </summary>
    [PublicAPI]
    public static class DocumentVisit
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

        /// <summary>
        ///
        /// </summary>
        /// <param name="cell">
        ///
        /// </param>
        /// <param name="revisionId">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        public static XElement ExecuteForTableRecursion([NotNull] XElement cell, int revisionId)
        {
            if (cell is null)
                throw new ArgumentNullException(nameof(cell));

            return Execute(cell, revisionId + 1);
        }

        /// <summary>
        /// Marshals content from the source document to be added into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="revisionId">
        /// The current revision number incremented by one.
        /// </param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        [NotNull]
        public static OpenXmlPackageVisitor VisitDoc([NotNull] this OpenXmlPackageVisitor subject, int revisionId)
        {
            if (subject is null)
                throw new ArgumentNullException(nameof(subject));

            Document document = subject.Document;

            return
                subject.With(
                    document.With(
                        Execute(document.Content, revisionId + 1)));
        }

        [Pure]
        [NotNull]
        private static XElement Execute([NotNull] XElement document, int revisionId)
        {
            if (document is null)
                throw new ArgumentNullException(nameof(document));

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
                    .RemoveByAll(W + "lang")
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
                    .SetTableStyles(revisionId)

                    // Remove elements used above, but not needed in the output.
                    .RemoveByAll(W + "u")
                    .RemoveByAllIfEmpty(W + "tcPr")
                    .RemoveByAllIfEmpty(W + "rPr")
                    .RemoveByAllIfEmpty(W + "pPr")
                    .RemoveByAllIfEmpty(W + "t")
                    .RemoveByAllIfEmpty(W + "r")
                    .RemoveByAll(x => x.Name == W + "p" && !x.HasElements && x.Parent?.Name != W + "tc")

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

                if (distinct.Any(x => (string) x.Attribute(W + "val") == "FootnoteReference"))
                    distinct = distinct.Where(x => (string) x.Attribute(W + "val") == "FootnoteReference");

                runProperties.AddFirst(distinct);
            }

            (int oldId, int newId)[] revisionMapping =
                source.Descendants()
                      .Where(x => Revisions.Contains(x.Name))
                      .Attributes(W + "id")
                      .Select(x => (int) x)
                      .OrderBy(x => x)
                      .Select(x => (oldId: x, newId: x + revisionId))
                      .OrderByDescending(x => x.oldId)
                      .ToArray();

            foreach (XName revision in Revisions)
            {
                foreach ((int oldId, int newId) in revisionMapping)
                {
                    source = source.ChangeXAttributeValues(revision, W + "id", oldId.ToString(), newId.ToString());
                }
            }

            source.Descendants(W + "sectPr").Attributes().Remove();

            source.Descendants(W + "p").Attributes().Remove();

            source.Descendants(W + "tr").Attributes().Remove();

            if (source.Element(W + "body")?.Elements().FirstOrDefault()?.Name == W + "sectPr")
                source.Element(W + "body")?.Elements().First().Remove();

            if (source.Element(W + "body")?.Elements().LastOrDefault()?.Name == W + "sectPr")
            {
                XElement sectionProperties = source.Element(W + "body")?.Elements().Last();

                XElement previous = sectionProperties.Previous();

                sectionProperties?.Remove();

                if (previous?.Name != W + "p")
                {
                    previous?.AddAfterSelf(new XElement(W + "p"));
                    previous = previous.Next();
                }

                if (!previous?.Elements(W + "pPr").Any() ?? false)
                    previous.AddFirst(new XElement(W + "pPr", sectionProperties));

                previous?.Element(W + "pPr")?.Add(sectionProperties);
            }

            IEnumerable<XElement> charts =
                source.Descendants(W + "drawing")
                      .Select(x => x.Ancestors(W + "p").FirstOrDefault())
                      .Where(x => x != null)
                      .ToArray();

            foreach (XElement item in charts)
            {
                item.RemoveByAll(W + "pStyle");

                if (!item.Elements(W + "pPr").Any())
                    item.AddFirst(new XElement(W + "pPr"));

                item.Element(W + "pPr")?
                    .AddFirst(
                        new XElement(W + "pStyle",
                            new XAttribute(W + "val",
                                "FiguresTablesSourceNote")));
            }

            return source;
        }
    }
}