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
    /// Add headers to a Word document.
    /// </summary>
    [PublicAPI]
    public static class AddHeadersExtensions
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
        /// Add headers to a Word document.
        /// </summary>
        public static void AddHeaders([NotNull] this DocxFilePath toFilePath, [NotNull] string title)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }
            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            // Modify [Content_Types].xml
            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");
            packageRelation.Descendants(T + "Override")
                           .Where(x => x.Attribute("PartName")?.Value.StartsWith("/word/header") ?? false)
                           .Remove();
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");

            // Modify document.xml.rels and grab the current header id number
            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Descendants(R + "Relationship")
                            .Where(x => x.Attribute("Target")?.Value.Contains("header") ?? false)
                            .Remove();
            int currentHeaderId = documentRelation.Elements().Attributes("Id").Select(x => int.Parse(x.Value.Substring(3))).DefaultIfEmpty(0).Max();
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            // Modify document.xml
            XElement document = toFilePath.ReadAsXml("word/document.xml");
            document.Descendants(W + "sectPr")
                    .Elements(W + "headerReference")
                    .Remove();
            document.WriteInto(toFilePath, "word/document.xml");
            
            // Add headers
            toFilePath.AddOddageHeader($"rId{++currentHeaderId}");
            toFilePath.AddEvenPageHeader($"rId{++currentHeaderId}", title);
        }

        private static void AddOddageHeader([NotNull] this DocxFilePath toFilePath, [NotNull] string headerId)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }
            if (headerId is null)
            {
                throw new ArgumentNullException(nameof(headerId));
            }

            XElement element = XElement.Parse(Resources.header1);
            element.WriteInto(toFilePath, "word/header1.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(R + "Relationship",
                    new XAttribute("Id", headerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute("Target", "header1.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml("word/document.xml");
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(W + "headerReference",
                        new XAttribute(W + "type", "default"),
                        new XAttribute(R + "id", headerId)));
            }
            document.WriteInto(toFilePath, "word/document.xml");

            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");
            packageRelation.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", "/word/header1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }

        private static void AddEvenPageHeader([NotNull] this DocxFilePath toFilePath, [NotNull]  string headerId, [NotNull] string title)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }
            if (headerId is null)
            {
                throw new ArgumentNullException(nameof(headerId));
            }
            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            XElement element = XElement.Parse(string.Format(Resources.header2, title));
            element.WriteInto(toFilePath, "word/header2.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(R + "Relationship",
                    new XAttribute("Id", headerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute("Target", "header2.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml("word/document.xml");
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(W + "headerReference",
                        new XAttribute(W + "type", "even"),
                        new XAttribute(R + "id", headerId)));
            }
            document.WriteInto(toFilePath, "word/document.xml");

            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");
            packageRelation.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", "/word/header2.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }
    }
}