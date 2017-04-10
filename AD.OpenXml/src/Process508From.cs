using System.Linq;
using System.Xml.Linq;
using AD.IO;
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
        /// <param name="fromFilePath">The document whose body is the target of the corrections.</param>
        /// <param name="toFilePath">The document to which the modified body is saved.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        public static void Process508From(this DocxFilePath toFilePath, DocxFilePath fromFilePath)
        {
            XElement element =
                fromFilePath.ReadAsXml()
                            .RemoveRsidAttributes()
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
                            //.AddLineBreakToHeadings()
                            .SetTableStyles()
                            .RemoveByAllIfEmpty(W + "rPr")
                            .RemoveByAllIfEmpty(W + "pPr")
                            .RemoveByAllIfEmpty(W + "t")
                            .RemoveByAllIfEmpty(W + "r")
                            .RemoveByAllIfEmpty(W + "p")
                            .RemoveByAll(W + "rFonts")
                            .RemoveByAll(W + "sz")
                            .RemoveByAll(W + "szCs")
                            .RemoveByAll(W + "u");
                            //.TransferCharts(fromFilePath, toFilePath);

            element.Descendants(W + "p").Attributes().Remove();
            element.Descendants(W + "tr").Attributes().Remove();
            element.Descendants().SelectMany(x => x.Elements().Where(y => y.Name == W + "pStyle").Skip(1)).Remove();
            element.Descendants(W + "hideMark").Remove();
            element.Descendants(W + "noWrap").Remove();
            element.Descendants(W + "rPr").Where(x => !x.HasElements).Remove();


            element.Element(W + "body")?
                   .Elements(W + "p")
                   .Elements(W + "pPr")
                   .Elements(W + "spacing")
                   .Remove();

            if (element.Element(W + "body")?.Elements().First().Name == W + "sectPr")
            {
                element.Element(W + "body")?.Elements().First().Remove();
            }
            element.Descendants(W + "hyperlink").Remove();

            element.WriteInto(toFilePath, "word/document.xml");
        }
    }
}