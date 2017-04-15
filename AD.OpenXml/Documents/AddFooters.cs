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
    /// Add footers to a Word document.
    /// </summary>
    [PublicAPI]
    public static class AddFootersExtensions
    {
        /// <summary>
        /// The namespace declared on the [Content_Types].xml
        /// </summary>
        [NotNull]
        private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of [Content_Types].xml
        /// </summary>
        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of document.xml.
        /// </summary>
        [NotNull]
        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;
        
        /// <summary>
        /// Add footers to a Word document.
        /// </summary>
        public static void AddFooters([NotNull] this DocxFilePath toFilePath)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }

            // Modify [Content_Types].xml
            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");

            packageRelation.Descendants(T + "Override")
                           .Where(x => x.Attribute("PartName")?.Value.StartsWith("/word/footer") ?? false)
                           .Remove();

            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");

            // Modify document.xml.rels and grab the current header id number
            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");

            documentRelation.Descendants(P + "Relationship")
                            .Where(x => x.Attribute("Target")?.Value.Contains("footer") ?? false)
                            .Remove();

            int currentFooterId = documentRelation.Elements().Attributes("Id").Select(x => int.Parse(x.Value.Substring(3))).DefaultIfEmpty(0).Max();

            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            // Modify document.xml
            XElement document = toFilePath.ReadAsXml("word/document.xml");

            document.Descendants(W + "sectPr")
                    .Elements(W + "footerReference")
                    .Remove();

            document.WriteInto(toFilePath, "word/document.xml");

            // Add footers
            toFilePath.AddEvenPageFooter($"rId{++currentFooterId}");
            toFilePath.AddOddPageFooter($"rId{++currentFooterId}");
        }

        private static void AddEvenPageFooter([NotNull] this DocxFilePath toFilePath, [NotNull] string footerId)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }
            if (footerId is null)
            {
                throw new ArgumentNullException(nameof(footerId));
            }

            XElement element = XElement.Parse(Resources.footer1);
            element.WriteInto(toFilePath, "word/footer1.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", footerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                    new XAttribute("Target", "footer1.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml("word/document.xml");
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.Add(
                    new XElement(W + "footerReference",
                        new XAttribute(W + "type", "even"),
                        new XAttribute(R + "id", footerId)));
            }
            document.WriteInto(toFilePath, "word/document.xml");

            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");
            packageRelation.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", "/word/footer1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }

        private static void AddOddPageFooter([NotNull] this DocxFilePath toFilePath, [NotNull] string footerId)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }
            if (footerId is null)
            {
                throw new ArgumentNullException(nameof(footerId));
            }

            XElement element = XElement.Parse(Resources.footer2);
            element.WriteInto(toFilePath, "word/footer2.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");

            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", footerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                    new XAttribute("Target", "footer2.xml")));

            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml("word/document.xml");

            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.Add(
                    new XElement(W + "footerReference",
                        new XAttribute(W + "type", "default"),
                        new XAttribute(R + "id", footerId)));
            }

            document.WriteInto(toFilePath, "word/document.xml");

            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");
            packageRelation.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", "/word/footer2.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }
    }
}