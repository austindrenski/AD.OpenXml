using System;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Properties;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Extension methods to add the styles part to a Word document.
    /// </summary>
    [PublicAPI]
    public static class AddStylesExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlPackageContentTypes;

        private static readonly XNamespace R = XNamespaces.OpenXmlPackageRelationships;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Adds the styles part to a Word document.
        /// </summary>
        /// <param name="toFilePath">The file to which styles are added.</param>
        /// <exception cref="ArgumentNullException"/>
        public static void AddStyles([NotNull] this DocxFilePath toFilePath)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }

            toFilePath.CreateStyles();
            toFilePath.CreateNumbering();
        }
        
        private static void CreateStyles([NotNull] this DocxFilePath toFilePath)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }

            XElement styles = toFilePath.ReadAsXml("word/styles.xml");
            styles.RemoveAll();

            styles.Add(XElement.Parse(Resources.DocDefaults));
            styles.Add(XElement.Parse(Resources.Normal));

            styles.Add(XElement.Parse(Resources.Appendix));
            styles.Add(XElement.Parse(Resources.Bibliography));
            styles.Add(XElement.Parse(Resources.Box));
            styles.Add(XElement.Parse(Resources.BoxCaption));
            styles.Add(XElement.Parse(Resources.BoxListBullet));
            styles.Add(XElement.Parse(Resources.BoxSourceNote));
            styles.Add(XElement.Parse(Resources.BlueTableBasic));
            styles.Add(XElement.Parse(Resources.CaptionFigure));
            styles.Add(XElement.Parse(Resources.CaptionTable));
            styles.Add(XElement.Parse(Resources.Emphasis));
            styles.Add(XElement.Parse(Resources.ExecutiveSummary1stParagraph));
            styles.Add(XElement.Parse(Resources.ExecutiveSummaryHighlights));
            styles.Add(XElement.Parse(Resources.ExecutiveSummarySidebars));
            styles.Add(XElement.Parse(Resources.FigureTableSourceNote));
            styles.Add(XElement.Parse(Resources.FootnoteReference));
            styles.Add(XElement.Parse(Resources.FootnoteText));
            styles.Add(XElement.Parse(Resources.Heading1));
            styles.Add(XElement.Parse(Resources.Heading2));
            styles.Add(XElement.Parse(Resources.Heading3));
            styles.Add(XElement.Parse(Resources.Heading4));
            styles.Add(XElement.Parse(Resources.Heading5));
            styles.Add(XElement.Parse(Resources.Heading6));
            styles.Add(XElement.Parse(Resources.Heading7));
            styles.Add(XElement.Parse(Resources.Heading8));
            styles.Add(XElement.Parse(Resources.Heading9));
            styles.Add(XElement.Parse(Resources.Hyperlink));
            styles.Add(XElement.Parse(Resources.ListBullet));
            styles.Add(XElement.Parse(Resources.PreHeading));
            styles.Add(XElement.Parse(Resources.Strong));
            styles.Add(XElement.Parse(Resources.StyleNotImplemented));
            styles.Add(XElement.Parse(Resources.Subscript));
            styles.Add(XElement.Parse(Resources.Superscript));
            styles.Add(XElement.Parse(Resources.TableOfFigures));
            styles.Add(XElement.Parse(Resources.TOC1));
            styles.Add(XElement.Parse(Resources.TOC2));
            styles.Add(XElement.Parse(Resources.TOC3));
            styles.Add(XElement.Parse(Resources.TOC4));
            styles.Add(XElement.Parse(Resources.TOCHeading));

            styles.WriteInto(toFilePath, "word/styles.xml");
        }

        private static void CreateNumbering([NotNull] this DocxFilePath toFilePath)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }

            XElement numbering = XElement.Parse(Resources.Numbering);
            numbering.WriteInto(toFilePath, "word/numbering.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");

            documentRelation.Descendants(R + "Relationship")
                            .Where(x => x.Attribute("Target")?.Value.Contains("numbering") ?? false)
                            .Remove();

            int documentId = documentRelation.Elements().Attributes("Id").Select(x => int.Parse(x.Value.Substring(3))).Max();

            documentRelation.Add(
                new XElement(R + "Relationship",
                    new XAttribute("Id", $"rId{++documentId}"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                    new XAttribute("Target", "numbering.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");

            packageRelation.Descendants(C + "Override")
                           .Where(x => x.Attribute("PartName")?.Value == "/word/numbering.xml")
                           .Remove();

            packageRelation.Add(
                new XElement(C + "Override",
                    new XAttribute("PartName", "/word/numbering.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }
    }
}