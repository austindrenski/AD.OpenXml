using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using AD.IO;
using AD.IO.Streams;
using AD.OpenXml.Structures;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    // TODO: write a HeaderVisit and document.
    /// <summary>
    /// Add headers to a Word document.
    /// </summary>
    [PublicAPI]
    public static class AddHeadersExtensions
    {
        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of document.xml.
        /// </summary>
        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// The content media type of an OpenXML header.
        /// </summary>
        [NotNull] private static readonly string HeaderContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";

        /// <summary>
        /// The schema type for an OpenXML header relationship.
        /// </summary>
        [NotNull] private static readonly string HeaderRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";

        /// <summary>
        ///
        /// </summary>
        [NotNull] private static readonly XElement Header1;

        /// <summary>
        ///
        /// </summary>
        [NotNull] private static readonly XElement Header2;

        /// <summary>
        ///
        /// </summary>
        static AddHeadersExtensions()
        {
            Assembly assembly = typeof(AddHeadersExtensions).GetTypeInfo().Assembly;

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Templates.Header1.xml"), Encoding.UTF8))
            {
                Header1 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Templates.Header2.xml"), Encoding.UTF8))
            {
                Header2 = XElement.Parse(reader.ReadToEnd());
            }
        }

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

            MemoryStream result = await stream.CopyPure();

            // Remove headers from [Content_Types].xml
            result =
                await result.ReadXml(ContentTypesInfo.Path)
                            .Recurse(x => (string) x.Attribute(ContentTypesInfo.Attributes.ContentType) != HeaderContentType)
                            .WriteIntoAsync(result, ContentTypesInfo.Path);

            // Remove headers from document.xml.rels
            result =
                await result.ReadXml(DocumentRelsInfo.Path)
                            .Recurse(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Type) != HeaderRelationshipType)
                            .WriteIntoAsync(result, DocumentRelsInfo.Path);

            // Remove headers from document.xml
            result =
                await result.ReadXml()
                            .Recurse(x => x.Name != W + "headerReference")
                            .WriteIntoAsync(result, "word/document.xml");

            // Store the current relationship id number
            int currentRelationshipId =
                result.ReadXml(DocumentRelsInfo.Path)
                      .Elements()
                      .Attributes("Id")
                      .Select(x => int.Parse(x.Value.Substring(3)))
                      .DefaultIfEmpty(0)
                      .Max();

            // Add headers
            result = await AddOddPageHeader(result, $"rId{++currentRelationshipId}");
            result = await AddEvenPageHeader(result, $"rId{++currentRelationshipId}", title);

            return result;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="headerId"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        [Pure]
        [NotNull]
        private static async Task<MemoryStream> AddOddPageHeader([NotNull] MemoryStream stream, [NotNull] string headerId)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (headerId is null)
            {
                throw new ArgumentNullException(nameof(headerId));
            }

            MemoryStream result = await Header2.WriteIntoAsync(stream, "word/header2.xml");

            XElement documentRelation = result.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(
                    DocumentRelsInfo.Elements.Relationship,
                    new XAttribute(
                        DocumentRelsInfo.Attributes.Id,
                        headerId),
                    new XAttribute(
                        DocumentRelsInfo.Attributes.Type,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute(
                        DocumentRelsInfo.Attributes.Target,
                        "header2.xml")));

            result = await documentRelation.WriteIntoAsync(result, DocumentRelsInfo.Path);

            XElement document = result.ReadXml();
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(
                        W + "headerReference",
                        new XAttribute(W + "type", "default"),
                        new XAttribute(R + "id", headerId)));
            }

            result = await document.WriteIntoAsync(result, "word/document.xml");

            XElement packageRelation = result.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(
                    ContentTypesInfo.Elements.Override,
                    new XAttribute(
                        ContentTypesInfo.Attributes.PartName,
                        "/word/header2.xml"),
                    new XAttribute(
                        ContentTypesInfo.Attributes.ContentType,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));

            return await packageRelation.WriteIntoAsync(result, ContentTypesInfo.Path);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="headerId"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        [Pure]
        [NotNull]
        private static async Task<MemoryStream> AddEvenPageHeader([NotNull] MemoryStream stream, [NotNull] string headerId, [NotNull] string title)
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

            MemoryStream result = await stream.CopyPure();

            XElement header1 = Header1.Clone();
            header1.Element(W + "p").Element(W + "r").Element(W + "t").Value = title;
            result = await header1.WriteIntoAsync(result, "word/header1.xml");

            XElement documentRelation = result.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(
                    DocumentRelsInfo.Elements.Relationship,
                    new XAttribute(
                        DocumentRelsInfo.Attributes.Id,
                        headerId),
                    new XAttribute(
                        DocumentRelsInfo.Attributes.Type,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute(
                        DocumentRelsInfo.Attributes.Target,
                        "header1.xml")));

            result = await documentRelation.WriteIntoAsync(result, DocumentRelsInfo.Path);

            XElement document = result.ReadXml();
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(
                        W + "headerReference",
                        new XAttribute(W + "type", "even"),
                        new XAttribute(R + "id", headerId)));
            }

            result = await document.WriteIntoAsync(result, "word/document.xml");

            XElement packageRelation = result.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(
                    ContentTypesInfo.Elements.Override,
                    new XAttribute(
                        ContentTypesInfo.Attributes.PartName,
                        "/word/header1.xml"),
                    new XAttribute(
                        ContentTypesInfo.Attributes.ContentType,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));

            return await packageRelation.WriteIntoAsync(result, ContentTypesInfo.Path);
        }
    }
}