using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Marshals content from the 'document.xml' file of a Word document as an idiomatic XML object.
    /// </summary>
    [PublicAPI]
    public static class MarshalContentFromExtensions
    {
        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Marshals content from the source document to be added into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        [NotNull]
        public static XElement MarshalContentFrom([NotNull] this DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            XElement source =
                file.ReadAsXml("word/document.xml")

                    // Remove editing attributes.
                    .RemoveRsidAttributes()

                    // Remove run properties from the paragraph scope.
                    .RemoveRunPropertiesFromParagraphProperties()

                    // Remove elements that should never exist in-line.
                    .RemoveByAll(W + "proofErr")
                    .RemoveByAll(W + "bookmarkStart")
                    .RemoveByAll(W + "bookmarkEnd")
                    .RemoveByAll(W + "tblPrEx")
                    .RemoveByAll(W + "spacing")
                    .RemoveByAll(W + "lang")
                    .RemoveByAll(W + "numPr")
                    .RemoveByAll(W + "hideMark")
                    .RemoveByAll(W + "noWrap")
                    .RemoveByAll(W + "rFonts")
                    .RemoveByAll(W + "sz")
                    .RemoveByAll(W + "szCs")
                    .RemoveByAll(W + "bCs")
                    .RemoveByAll(W + "iCs")
                    .RemoveByAll(W + "color")
                    .RemoveByAll(W + "lastRenderedPageBreak")
                    .RemoveByAll(W + "keepNext")
                    .RemoveByAll(W + "noProof")

                    // Remove elements that should almost never exist.
                    .RemoveByAll(x => x.Name.Equals(W + "br") && (x.Attribute(W + "type")?.Value.Equals("page", StringComparison.OrdinalIgnoreCase) ?? false))
                    .RemoveByAll(x => x.Name.Equals(W + "pStyle") && (x.Attribute(W + "val")?.Value.Equals("BodyTextSSFinal", StringComparison.OrdinalIgnoreCase) ?? false))
                    .RemoveByAll(x => x.Name.Equals(W + "pStyle") && (x.Attribute(W + "val")?.Value.Equals("Default", StringComparison.OrdinalIgnoreCase) ?? false))
                    .RemoveByAll(x => x.Name.Equals(W + "jc") && !x.Ancestors(W + "table").Any())

                    // Alter bold, italic, and underline elements.
                    .ChangeBoldToStrong()
                    .ChangeItalicToEmphasis()
                    .ChangeUnderlineToTableCaption()
                    .ChangeUnderlineToFigureCaption()
                    .ChangeUnderlineToSourceNote()
                    .ChangeSuperscriptToReference()

                    // Mark insert requests for the production team.
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

                    // Tidy up the XML for review.
                    .MergeRuns();

            // Set page size.
            foreach (XElement pageSize in source.Descendants(W + "pgSz"))
            {
                pageSize.Element(W + "pgSz")?.SetAttributeValue(W + "w", "12240");
                pageSize.Element(W + "pgSz")?.SetAttributeValue(W + "h", "15840");
            }

            // There shouldn't be section properties without orientations.
            foreach (XElement sectionProperties in source.Descendants(W + "sectPr").Where(x => !x.Descendants().Attributes(W + "orient").Any()).ToArray())
            {
                sectionProperties.Element(W + "pgSz")?.SetAttributeValue(W + "orient", "portrait");
            }

            // There shouldn't be section properties type=nextPage.
            foreach (XElement sectionProperties in source.Descendants(W + "sectPr").Where(x => x.Element(W + "type")?.Value != "nextPage").ToArray())
            {
                XElement type = sectionProperties.Element(W + "type");
                if (type is null)
                {
                    sectionProperties.Add(new XElement(W + "type"));
                }
                sectionProperties.Element(W + "type")?.SetAttributeValue(W + "val", "nextPage");
            }

            // There shouldn't be section properties in paragraph properties.
            foreach (XElement sectionProperties in source.Descendants(W + "sectPr").Where(x => x.Ancestors(W + "pPr").Any()).ToArray())
            {
                XElement ancestorParagraph = sectionProperties.Ancestors(W + "p").FirstOrDefault();
                if (ancestorParagraph is null)
                {
                    continue;
                }
                sectionProperties.Remove();
                ancestorParagraph.AddAfterSelf(sectionProperties);
            }
            
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
                IEnumerable<XElement> distinct = styles.Distinct().ToArray();
                if (distinct.Any(x => x.Attribute(W + "val")?.Value.Equals("FootnoteReference") ?? false))
                {
                    distinct = distinct.Where(x => x.Attribute(W + "val")?.Value.Equals("FootnoteReference") ?? false);
                }
                runProperties.AddFirst(distinct);
            }

            source.Descendants(W + "sectPr").Attributes().Remove();

            source.Descendants(W + "p").Attributes().Remove();

            source.Descendants(W + "tr").Attributes().Remove();

            if (source.Element(W + "body")?.Elements().First().Name == W + "sectPr")
            {
                source.Element(W + "body")?.Elements().First().Remove();
            }

            source.Descendants(W + "hyperlink").Remove();

            return source;
        }
    }
}