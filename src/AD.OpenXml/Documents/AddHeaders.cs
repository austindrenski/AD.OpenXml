using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using AD.IO;
using AD.IO.Paths;
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
        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of [Content_Types].xml
        /// </summary>
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of document.xml.
        /// </summary>
        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Add headers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static async Task<MemoryStream> AddHeaders([NotNull] this Task<MemoryStream> stream, [NotNull] string title)
        {
            return await AddHeaders(await stream, title);
        }

        /// <summary>
        /// Add headers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static async Task<MemoryStream> AddHeaders([NotNull] this MemoryStream stream, [NotNull] string title)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            MemoryStream result = new MemoryStream();
            await stream.CopyToAsync(result);

            // Modify [Content_Types].xml
            XElement packageRelation = result.ReadAsXml("[Content_Types].xml");
            packageRelation.Descendants(T + "Override")
                           .Where(x => x.Attribute("PartName")?.Value.StartsWith("/word/header") ?? false)
                           .Remove();
            result = await packageRelation.WriteInto(result, "[Content_Types].xml");

            // Modify document.xml.rels and grab the current header id number
            XElement documentRelation = result.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Descendants(R + "Relationship")
                            .Where(x => x.Attribute("Target")?.Value.Contains("header") ?? false)
                            .Remove();
            int currentHeaderId = documentRelation.Elements().Attributes("Id").Select(x => int.Parse(x.Value.Substring(3))).DefaultIfEmpty(0).Max();
            result = await documentRelation.WriteInto(result, "word/_rels/document.xml.rels");

            // Modify document.xml
            XElement document = result.ReadAsXml("document.xml");
            document.Descendants(W + "sectPr")
                    .Elements(W + "headerReference")
                    .Remove();
            result = await document.WriteInto(result, "word/document.xml");

            // Add headers
            result = await result.AddOddPageHeader($"rId{++currentHeaderId}");
            result = await result.AddEvenPageHeader($"rId{++currentHeaderId}", title);

            return result;
        }

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
            XElement document = toFilePath.ReadAsXml();
            document.Descendants(W + "sectPr")
                    .Elements(W + "headerReference")
                    .Remove();
            document.WriteInto(toFilePath, "word/document.xml");

            // Add headers
            toFilePath.AddOddPageHeader($"rId{++currentHeaderId}");
            toFilePath.AddEvenPageHeader($"rId{++currentHeaderId}", title);
        }

        private static async Task<MemoryStream> AddOddPageHeader([NotNull] this MemoryStream stream, [NotNull] string headerId)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            if (headerId is null)
            {
                throw new ArgumentNullException(nameof(headerId));
            }

            MemoryStream result = new MemoryStream();
            await stream.CopyToAsync(result);

            XElement element = XElement.Parse(Resources.header2);
            result = await element.WriteInto(result, "word/header2.xml");

            XElement documentRelation = result.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", headerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute("Target", "header2.xml")));
            result = await documentRelation.WriteInto(result, "word/_rels/document.xml.rels");

            XElement document = result.ReadAsXml("document.xml");
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(W + "headerReference",
                        new XAttribute(W + "type", "default"),
                        new XAttribute(R + "id", headerId)));
            }

            result = await document.WriteInto(result, "word/document.xml");

            XElement packageRelation = result.ReadAsXml("[Content_Types].xml");
            packageRelation.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", "/word/header2.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
            result = await packageRelation.WriteInto(result, "[Content_Types].xml");

            return result;
        }


        private static void AddOddPageHeader([NotNull] this DocxFilePath toFilePath, [NotNull] string headerId)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }
            if (headerId is null)
            {
                throw new ArgumentNullException(nameof(headerId));
            }

            XElement element = XElement.Parse(Resources.header2);
            element.WriteInto(toFilePath, "word/header2.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", headerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute("Target", "header2.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml();
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
                    new XAttribute("PartName", "/word/header2.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }

        private static async Task<MemoryStream> AddEvenPageHeader([NotNull] this MemoryStream stream, [NotNull] string headerId, [NotNull] string title)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }
            if (headerId is null)
            {
                throw new ArgumentNullException(nameof(headerId));
            }
            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            MemoryStream result = new MemoryStream();
            await stream.CopyToAsync(result);

            XElement element = XElement.Parse(string.Format(Resources.header1, title));
            result = await element.WriteInto(result, "word/header1.xml");

            XElement documentRelation = result.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", headerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute("Target", "header1.xml")));
            result = await documentRelation.WriteInto(result, "word/_rels/document.xml.rels");

            XElement document = result.ReadAsXml("document.xml");
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(W + "headerReference",
                        new XAttribute(W + "type", "even"),
                        new XAttribute(R + "id", headerId)));
            }

            result = await document.WriteInto(result, "word/document.xml");

            XElement packageRelation = result.ReadAsXml("[Content_Types].xml");
            packageRelation.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", "/word/header1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
            result = await packageRelation.WriteInto(result, "[Content_Types].xml");

            return result;
        }


        private static void AddEvenPageHeader([NotNull] this DocxFilePath toFilePath, [NotNull] string headerId, [NotNull] string title)
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

            XElement element = XElement.Parse(string.Format(Resources.header1, title));
            element.WriteInto(toFilePath, "word/header1.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", headerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute("Target", "header1.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml();
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
                    new XAttribute("PartName", "/word/header1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }
    }
}