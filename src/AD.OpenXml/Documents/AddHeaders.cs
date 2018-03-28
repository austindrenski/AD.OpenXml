using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
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

        [Pure]
        [NotNull]
        private static ZipArchive Clone([NotNull] ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            ZipArchive writer = new ZipArchive(new MemoryStream(), ZipArchiveMode.Update);

            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                using (Stream readStream = entry.Open())
                {
                    using (Stream writeStream = writer.CreateEntry(entry.FullName).Open())
                    {
                        readStream.CopyTo(writeStream);
                    }
                }
            }

            return writer;
        }

        [Pure]
        [NotNull]
        private static MemoryStream CloneToStream([NotNull] ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            MemoryStream ms = new MemoryStream();

            using (ZipArchive writer = new ZipArchive(ms, ZipArchiveMode.Update, true))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    using (Stream readStream = entry.Open())
                    {
                        using (Stream writeStream = writer.CreateEntry(entry.FullName).Open())
                        {
                            readStream.CopyTo(writeStream);
                        }
                    }
                }
            }

            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }

        /// <summary>
        /// Add headers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        // TODO: return ZipArchive
        public static MemoryStream AddHeaders([NotNull] this ZipArchive archive, [NotNull] string title)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            ZipArchive result = Clone(archive);

            // Remove headers from [Content_Types].xml
            result.GetEntry(ContentTypesInfo.Path).Delete();
            using (Stream stream = result.CreateEntry(ContentTypesInfo.Path).Open())
            {
                archive.ReadXml(ContentTypesInfo.Path)
                       .Recurse(x => (string) x.Attribute(ContentTypesInfo.Attributes.ContentType) != HeaderContentType)
                       .Save(stream);
            }

            // Remove headers from document.xml.rels
            result.GetEntry(DocumentRelsInfo.Path).Delete();
            using (Stream stream = result.CreateEntry(DocumentRelsInfo.Path).Open())
            {
                archive.ReadXml(DocumentRelsInfo.Path)
                       .Recurse(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Type) != HeaderRelationshipType)
                       .Save(stream);
            }

            // Remove headers from document.xml
            result.GetEntry("word/document.xml").Delete();
            using (Stream stream = result.CreateEntry("word/document.xml").Open())
            {
                archive.ReadXml()
                       .Recurse(x => x.Name != W + "headerReference")
                       .Save(stream);
            }

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

            return CloneToStream(result);
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
            archive.GetEntry("word/header2.xml")?.Delete();
            using (Stream stream = archive.CreateEntry("word/header2.xml").Open())
            {
                stream.SetLength(0);
                Header2.Save(stream);
            }

            XElement documentRelation = archive.ReadXml(DocumentRelsInfo.Path);

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
                    new XAttribute(
                        ContentTypesInfo.Attributes.PartName,
                        "/word/header2.xml"),
                    new XAttribute(
                        ContentTypesInfo.Attributes.ContentType,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));

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

            XElement header1 = Header1.Clone();
            header1.Element(W + "p").Element(W + "r").Element(W + "t").Value = title;

            archive.GetEntry("word/header1.xml")?.Delete();
            using (Stream stream = archive.CreateEntry("word/header1.xml").Open())
            {
                stream.SetLength(0);
                header1.Save(stream);
            }

            XElement documentRelation = archive.ReadXml(DocumentRelsInfo.Path);

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
                    new XAttribute(
                        ContentTypesInfo.Attributes.PartName,
                        "/word/header1.xml"),
                    new XAttribute(
                        ContentTypesInfo.Attributes.ContentType,
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));

            using (Stream stream = archive.GetEntry(ContentTypesInfo.Path).Open())
            {
                stream.SetLength(0);
                packageRelation.Save(stream);
            }
        }
    }
}