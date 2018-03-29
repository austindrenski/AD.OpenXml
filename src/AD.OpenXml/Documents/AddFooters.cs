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
    // TODO: write a FooterVisit and document.
    /// <summary>
    /// Add footers to a Word document.
    /// </summary>
    [PublicAPI]
    public static class AddFootersExtensions
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
        /// The content media type of an OpenXML footer.
        /// </summary>
        [NotNull] private static readonly string FooterContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

        /// <summary>
        /// The schema type for an OpenXML footer relationship.
        /// </summary>
        [NotNull] private static readonly string FooterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

        [NotNull] private static readonly XElement Footer1;
        [NotNull] private static readonly XElement Footer2;

        /// <summary>
        ///
        /// </summary>
        static AddFootersExtensions()
        {
            Assembly assembly = typeof(AddFootersExtensions).GetTypeInfo().Assembly;

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Templates.Footer1.xml"), Encoding.UTF8))
            {
                Footer1 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Templates.Footer2.xml"), Encoding.UTF8))
            {
                Footer2 = XElement.Parse(reader.ReadToEnd());
            }
        }

        /// <summary>
        /// Add footers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static ZipArchive AddFooters([NotNull] this ZipArchive archive, [NotNull] string publisher, [NotNull] string website)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            ZipArchive result =
                archive.With(
                    // Remove footers from [Content_Types].xml
                    (
                        ContentTypesInfo.Path,
                        z => z.ReadXml(ContentTypesInfo.Path)
                              .Recurse(x => (string) x.Attribute(ContentTypesInfo.Attributes.ContentType) != FooterContentType)
                    ),
                    // Remove footers from document.xml.rels
                    (
                        DocumentRelsInfo.Path,
                        z => z.ReadXml(DocumentRelsInfo.Path)
                              .Recurse(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Type) != FooterRelationshipType)
                    ),
                    // Remove footers from document.xml
                    (
                        "word/document.xml",
                        z => z.ReadXml()
                              .Recurse(x => x.Name != W + "footerReference")
                    ));

            // Store the current relationship id number
            int currentRelationshipId =
                result.ReadXml(DocumentRelsInfo.Path)
                      .Elements()
                      .Attributes("Id")
                      .Select(x => int.Parse(x.Value.Substring(3)))
                      .DefaultIfEmpty(0)
                      .Max();

            // Add footers
            AddEvenPageFooter(result, $"rId{++currentRelationshipId}", website);
            AddOddPageFooter(result, $"rId{++currentRelationshipId}", publisher);

            return result;
        }

        private static void AddOddPageFooter([NotNull] ZipArchive archive, [NotNull] string footerId, [NotNull] string publisher)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            if (footerId is null)
            {
                throw new ArgumentNullException(nameof(footerId));
            }

            if (publisher is null)
            {
                throw new ArgumentNullException(nameof(publisher));
            }

            XElement footer1 = Footer1.Clone();
            footer1.Element(W + "p").Elements(W + "r").First().Element(W + "t").Value = publisher;

            archive.GetEntry("word/footer1.xml")?.Delete();
            using (Stream stream = archive.CreateEntry("word/footer1.xml").Open())
            {
                stream.SetLength(0);
                footer1.Save(stream);
            }

            XElement documentRelation = archive.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", footerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                    new XAttribute("Target", "footer1.xml")));

            archive.GetEntry(DocumentRelsInfo.Path)?.Delete();
            using (Stream stream = archive.CreateEntry(DocumentRelsInfo.Path).Open())
            {
                stream.SetLength(0);
                documentRelation.Save(stream);
            }

            XElement document = archive.ReadXml();

            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                if (sectionProperties.Elements(W + "headerReference").Any())
                {
                    sectionProperties.Elements(W + "headerReference")
                                     .Last()
                                     .AddAfterSelf(
                                         new XElement(
                                             W + "footerReference",
                                             new XAttribute(W + "type", "default"),
                                             new XAttribute(R + "id", footerId)));
                }
                else
                {
                    sectionProperties.Add(
                        new XElement(
                            W + "footerReference",
                            new XAttribute(W + "type", "default"),
                            new XAttribute(R + "id", footerId)));
                }
            }

            archive.GetEntry("word/document.xml")?.Delete();
            using (Stream stream = archive.CreateEntry("word/document.xml").Open())
            {
                stream.SetLength(0);
                document.Save(stream);
            }

            XElement packageRelation = archive.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", "/word/footer1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));

            archive.GetEntry(ContentTypesInfo.Path)?.Delete();
            using (Stream stream = archive.CreateEntry(ContentTypesInfo.Path).Open())
            {
                stream.SetLength(0);
                packageRelation.Save(stream);
            }
        }

        private static void AddEvenPageFooter([NotNull] ZipArchive archive, [NotNull] string footerId, [NotNull] string website)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            if (footerId is null)
            {
                throw new ArgumentNullException(nameof(footerId));
            }

            if (website is null)
            {
                throw new ArgumentNullException(nameof(website));
            }

            XElement footer2 = Footer2.Clone();
            footer2.Element(W + "p").Elements(W + "r").Last().Element(W + "t").Value = website;

            archive.GetEntry("word/footer2.xml")?.Delete();
            using (Stream stream = archive.CreateEntry("word/footer2.xml").Open())
            {
                stream.SetLength(0);
                footer2.Save(stream);
            }

            XElement documentRelation = archive.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", footerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                    new XAttribute("Target", "footer2.xml")));

            archive.GetEntry(DocumentRelsInfo.Path)?.Delete();
            using (Stream stream = archive.CreateEntry(DocumentRelsInfo.Path).Open())
            {
                stream.SetLength(0);
                documentRelation.Save(stream);
            }

            XElement document = archive.ReadXml();

            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                if (sectionProperties.Elements(W + "headerReference").Any())
                {
                    sectionProperties.Elements(W + "headerReference")
                                     .Last()
                                     .AddAfterSelf(
                                         new XElement(
                                             W + "footerReference",
                                             new XAttribute(W + "type", "even"),
                                             new XAttribute(R + "id", footerId)));
                }
                else
                {
                    sectionProperties.Add(
                        new XElement(
                            W + "footerReference",
                            new XAttribute(W + "type", "even"),
                            new XAttribute(R + "id", footerId)));
                }
            }

            archive.GetEntry("word/document.xml")?.Delete();
            using (Stream stream = archive.CreateEntry("word/document.xml").Open())
            {
                stream.SetLength(0);
                document.Save(stream);
            }

            XElement packageRelation = archive.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", "/word/footer2.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));

            archive.GetEntry(ContentTypesInfo.Path)?.Delete();
            using (Stream stream = archive.CreateEntry(ContentTypesInfo.Path).Open())
            {
                stream.SetLength(0);
                packageRelation.Save(stream);
            }
        }
    }
}