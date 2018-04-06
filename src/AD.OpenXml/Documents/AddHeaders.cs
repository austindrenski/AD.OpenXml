using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
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
        /// The XML declaration.
        /// </summary>
        [NotNull] private static readonly XDeclaration Declaration = new XDeclaration("1.0", "utf-8", "yes");

        /// <summary>
        /// Add headers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static ZipArchive AddHeaders([NotNull] this ZipArchive archive, [NotNull] string title)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            ZipArchive result =
                archive.With(
                    // Remove headers from [Content_Types].xml
                    (
                        ContentTypesInfo.Path,
                        z => z.ReadXml(ContentTypesInfo.Path)
                              .Recurse(x => (string) x.Attribute(ContentTypesInfo.Attributes.ContentType) != HeaderContentType)
                    ),
                    // Remove headers from document.xml.rels
                    (
                        DocumentRelsInfo.Path,
                        z => z.ReadXml(DocumentRelsInfo.Path)
                              .Recurse(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Type) != HeaderRelationshipType)
                    ),
                    // Remove headers from document.xml
                    (
                        "word/document.xml",
                        z => z.ReadXml()
                              .Recurse(x => x.Name != W + "headerReference")
                    ));

            // Store the current relationship id number
            int currentRelationshipId =
                result.ReadXml(DocumentRelsInfo.Path)
                      .Elements()
                      .Attributes("Id")
                      .Select(x => int.Parse(x.Value.Substring(3)))
                      .DefaultIfEmpty(0)
                      .Max();

            // Add headers
            AddOddPageHeader(result, $"rId{++currentRelationshipId}");
            AddEvenPageHeader(result, $"rId{++currentRelationshipId}", title);

            return result;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="archive"></param>
        /// <param name="headerId"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private static void AddOddPageHeader([NotNull] ZipArchive archive, [NotNull] string headerId)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            if (headerId is null)
            {
                throw new ArgumentNullException(nameof(headerId));
            }

            // Remove headers from document.xml
            using (Stream stream = archive.GetEntry("word/header2.xml")?.Open() ??
                                   archive.CreateEntry("word/header2.xml").Open())
            {
                stream.SetLength(0);
                Header2().Save(stream);
            }

            XElement documentRelation = archive.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(
                    DocumentRelsInfo.Elements.Relationship,
                    new XAttribute(DocumentRelsInfo.Attributes.Id, headerId),
                    new XAttribute(DocumentRelsInfo.Attributes.Type, HeaderRelationshipType),
                    new XAttribute(DocumentRelsInfo.Attributes.Target, "header2.xml")));

            using (Stream stream = archive.GetEntry(DocumentRelsInfo.Path).Open())
            {
                stream.SetLength(0);
                documentRelation.Save(stream);
            }

            XElement document = archive.ReadXml();
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(
                        W + "headerReference",
                        new XAttribute(W + "type", "default"),
                        new XAttribute(R + "id", headerId)));
            }

            using (Stream stream = archive.GetEntry("word/document.xml").Open())
            {
                stream.SetLength(0);
                document.Save(stream);
            }

            XElement packageRelation = archive.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(
                    ContentTypesInfo.Elements.Override,
                    new XAttribute(ContentTypesInfo.Attributes.PartName, "/word/header2.xml"),
                    new XAttribute(ContentTypesInfo.Attributes.ContentType, HeaderContentType)));

            using (Stream stream = archive.GetEntry(ContentTypesInfo.Path).Open())
            {
                stream.SetLength(0);
                packageRelation.Save(stream);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="archive"></param>
        /// <param name="headerId"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private static void AddEvenPageHeader([NotNull] ZipArchive archive, [NotNull] string headerId, [NotNull] string title)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            if (headerId is null)
            {
                throw new ArgumentNullException(nameof(headerId));
            }

            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            XDocument header1 = Header1(title);

            using (Stream stream = archive.GetEntry("word/header1.xml")?.Open() ??
                                   archive.CreateEntry("word/header1.xml").Open())
            {
                stream.SetLength(0);
                header1.Save(stream);
            }

            XElement documentRelation = archive.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(
                    DocumentRelsInfo.Elements.Relationship,
                    new XAttribute(DocumentRelsInfo.Attributes.Id, headerId),
                    new XAttribute(DocumentRelsInfo.Attributes.Type, HeaderRelationshipType),
                    new XAttribute(DocumentRelsInfo.Attributes.Target, "header1.xml")));

            using (Stream stream = archive.GetEntry(DocumentRelsInfo.Path).Open())
            {
                stream.SetLength(0);
                documentRelation.Save(stream);
            }

            XElement document = archive.ReadXml();
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(
                        W + "headerReference",
                        new XAttribute(W + "type", "even"),
                        new XAttribute(R + "id", headerId)));
            }

            using (Stream stream = archive.GetEntry("word/document.xml").Open())
            {
                stream.SetLength(0);
                document.Save(stream);
            }

            XElement packageRelation = archive.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(
                    ContentTypesInfo.Elements.Override,
                    new XAttribute(ContentTypesInfo.Attributes.PartName, "/word/header1.xml"),
                    new XAttribute(ContentTypesInfo.Attributes.ContentType, HeaderContentType)));

            using (Stream stream = archive.GetEntry(ContentTypesInfo.Path).Open())
            {
                stream.SetLength(0);
                packageRelation.Save(stream);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="title">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static XDocument Header1([NotNull] string title)
        {
            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            return
                new XDocument(
                    Declaration,
                    new XElement(
                        W + "hdr",
                        new XElement(
                            W + "p",
                            new XElement(
                                W + "pPr",
                                new XElement(
                                    W + "pStyle",
                                    new XAttribute(W + "val", "Header")),
                                new XElement(
                                    W + "jc",
                                    new XAttribute(W + "val", "left"))),
                            new XElement(
                                W + "r",
                                new XElement(
                                    W + "t",
                                    new XText(title))))));
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static XDocument Header2()
        {
            return
                new XDocument(
                    Declaration,
                    new XElement(
                        W + "hdr",
                        new XElement(
                            W + "p",
                            new XElement(
                                W + "pPr",
                                new XElement(
                                    W + "pStyle",
                                    new XAttribute(W + "val", "Header")),
                                new XElement(
                                    W + "jc",
                                    new XAttribute(W + "val", "right"))),
                            new XElement(
                                W + "r",
                                new XElement(
                                    W + "t",
                                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                                    new XText("Chapter "))),
                            new XElement(
                                W + "fldSimple",
                                new XAttribute(W + "instr", " STYLEREF  \"heading 1\" \\s \\* MERGEFORMAT "),
                                new XElement(
                                    W + "r",
                                    new XElement(
                                        W + "t",
                                        new XText(string.Empty)))),
                            new XElement(
                                W + "r",
                                new XElement(
                                    W + "t",
                                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                                    new XText(": "))),
                            new XElement(
                                W + "fldSimple",
                                new XAttribute(W + "instr", " STYLEREF  \"heading 1\" \\* MERGEFORMAT "),
                                new XElement(
                                    W + "r",
                                    new XElement(
                                        W + "t",
                                        new XText(string.Empty)))))));
        }
    }
}