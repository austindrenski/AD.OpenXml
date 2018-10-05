using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public static class FootnoteVisit
    {
        [NotNull] static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] static readonly IEnumerable<XName> Revisions =
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
        [NotNull]
        public static OpenXmlPackageVisitor VisitFootnotes([NotNull] this OpenXmlPackageVisitor subject, int footnoteId, int revisionId)
        {
            if (subject is null)
                throw new ArgumentNullException(nameof(subject));

            (XElement document, XElement footnotes) =
                Execute(
                    subject.Footnotes.Content,
                    subject.Document.Content,
                    footnoteId + 1,
                    revisionId + 1);

            return
                subject.With(
                    subject.Document.With(document),
                    subject.Footnotes.With(footnotes));
        }

        [Pure]
        static (XElement Document, XElement Footnotes) Execute([NotNull] XElement footnotes, [NotNull] XElement document, int footnoteId, int revisionId)
        {
            ReportVisitor visitor = new ReportVisitor();

            if (!(visitor.Visit(footnotes) is XElement visited))
                throw new ArgumentException("This should never be thrown.");

            XElement modifiedFootnotes =
                visited

                   .RemoveByAll(W + "commentRangeStart")
                   .RemoveByAll(W + "commentRangeEnd")
                   .RemoveByAll(W + "commentReference")
                   .RemoveByAll(x => (string) x.Attribute(W + "val") == "CommentReference")

                    // Remove elements that should almost never exist.
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
                    distinct = distinct.Where(x => (string) x.Attribute(W + "val") == "FootnoteReference");

                runProperties.AddFirst(distinct);
            }

            (int oldId, int newId)[] footnoteMapping =
                modifiedFootnotes.Elements(W + "footnote")
                                 .Select(x => (int) x.Attribute(W + "id"))
                                 .OrderBy(x => x)
                                 .Select((x, i) => (oldId: x, newId: i + footnoteId))
                                 .OrderByDescending(x => x.oldId)
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
                                 .OrderByDescending(x => x.oldId)
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