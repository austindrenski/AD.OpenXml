using System;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Extension methods to support 508-compliance.
    /// </summary>
    [PublicAPI]
    public static class Process508FromExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Performs a variety of 508-compliance corrections.
        /// This method works on the existing <see cref="XElement"/> and returns a reference to it for a fluent syntax.
        /// </summary>
        /// <param name="source">The document whose body is the target of the corrections.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="ArgumentException"/>
        /// <exception cref="ArgumentNullException"/>
        public static XElement Process508From([NotNull] this XElement source)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            XElement element =
                source.RemoveRsidAttributes()
                      .RemoveRunPropertiesFromParagraphProperties()
                      .RemoveByAll(W + "proofErr")
                      .RemoveByAll(W + "bookmarkStart")
                      .RemoveByAll(W + "bookmarkEnd")
                      .MergeRuns()
                      .ChangeBoldToStrong()
                      .ChangeItalicToEmphasis()
                      .ChangeUnderlineToTableCaption()
                      .ChangeUnderlineToFigureCaption()
                      .ChangeUnderlineToSourceNote()
                      .ChangeSuperscriptToReference()
                      .HighlightInsertRequests()
                      .AddLineBreakToHeadings()
                      .SetTableStyles()
                      .RemoveByAll(W + "rFonts")
                      .RemoveByAll(W + "sz")
                      .RemoveByAll(W + "szCs")
                      .RemoveByAll(W + "u")
                      .RemoveByAllIfEmpty(W + "rPr")
                      .RemoveByAllIfEmpty(W + "pPr")
                      .RemoveByAllIfEmpty(W + "t")
                      .RemoveByAllIfEmpty(W + "r")
                      .RemoveByAllIfEmpty(W + "p");

            element.Descendants(W + "rPr")
                   .Where(
                       x =>
                           x.Elements(W + "rStyle")
                            .Attributes(W + "val")
                            .Any(y => y.Value.Equals("FootnoteReference")))
                   .SelectMany(
                       x =>
                           x.Descendants()
                            .Where(y => !y.Attribute(W + "val")?.Value.Equals("FootnoteReference") ?? false))
                   .Remove();

            element.Descendants(W + "p").Attributes().Remove();
            element.Descendants(W + "tr").Attributes().Remove();
            element.Descendants(W + "hideMark").Remove();
            element.Descendants(W + "noWrap").Remove();
            element.Descendants(W + "pPr").Where(x => !x.HasElements).Remove();
            element.Descendants(W + "rPr").Where(x => !x.HasElements).Remove();
            element.Descendants(W + "spacing").Remove();
            
            if (element.Element(W + "body")?.Elements().First().Name == W + "sectPr")
            {
                element.Element(W + "body")?.Elements().First().Remove();
            }

            element.Descendants(W + "hyperlink").Remove();

            return element;
        }
    }
}