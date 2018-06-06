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
        /// The XML declaration.
        /// </summary>
        [NotNull] private static readonly XDeclaration Declaration = new XDeclaration("1.0", "utf-8", "yes");

        /// <summary>
        /// The content media type of an OpenXML footer.
        /// </summary>
        [NotNull] private static readonly string FooterContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

        /// <summary>
        /// The schema type for an OpenXML footer relationship.
        /// </summary>
        [NotNull] private static readonly string FooterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

        /// <summary>
        /// Add footers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static ZipArchive AddFooters([NotNull] this ZipArchive archive, [NotNull] string publisher, [NotNull] string website)
        {
            if (archive is null)
                throw new ArgumentNullException(nameof(archive));

            if (publisher is null)
                throw new ArgumentNullException(nameof(publisher));

            if (website is null)
                throw new ArgumentNullException(nameof(website));

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
                throw new ArgumentNullException(nameof(archive));

            if (footerId is null)
                throw new ArgumentNullException(nameof(footerId));

            if (publisher is null)
                throw new ArgumentNullException(nameof(publisher));

            XDocument footer1 = Footer1(publisher);

            using (Stream stream = archive.GetEntry("word/footer1.xml")?.Open() ??
                                   archive.CreateEntry("word/footer1.xml").Open())
            {
                stream.SetLength(0);
                footer1.Save(stream);
            }

            XElement documentRelation = archive.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(
                    P + "Relationship",
                    new XAttribute("Id", footerId),
                    new XAttribute("Type", FooterRelationshipType),
                    new XAttribute("Target", "footer1.xml")));

            using (Stream stream = archive.GetEntry(DocumentRelsInfo.Path)?.Open() ??
                                   archive.CreateEntry(DocumentRelsInfo.Path).Open())
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

            using (Stream stream = archive.GetEntry("word/document.xml")?.Open() ??
                                   archive.CreateEntry("word/document.xml").Open())
            {
                stream.SetLength(0);
                document.Save(stream);
            }

            XElement packageRelation = archive.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(
                    T + "Override",
                    new XAttribute("PartName", "/word/footer1.xml"),
                    new XAttribute("ContentType", FooterContentType)));

            using (Stream stream = archive.GetEntry(ContentTypesInfo.Path)?.Open() ??
                                   archive.CreateEntry(ContentTypesInfo.Path).Open())
            {
                stream.SetLength(0);
                packageRelation.Save(stream);
            }
        }

        private static void AddEvenPageFooter([NotNull] ZipArchive archive, [NotNull] string footerId, [NotNull] string website)
        {
            if (archive is null)
                throw new ArgumentNullException(nameof(archive));

            if (footerId is null)
                throw new ArgumentNullException(nameof(footerId));

            if (website is null)
                throw new ArgumentNullException(nameof(website));

            XDocument footer2 = Footer2(website);

            using (Stream stream = archive.GetEntry("word/footer2.xml")?.Open() ??
                                   archive.CreateEntry("word/footer2.xml").Open())
            {
                stream.SetLength(0);
                footer2.Save(stream);
            }

            XElement documentRelation = archive.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(
                    P + "Relationship",
                    new XAttribute("Id", footerId),
                    new XAttribute("Type", FooterRelationshipType),
                    new XAttribute("Target", "footer2.xml")));

            using (Stream stream = archive.GetEntry(DocumentRelsInfo.Path)?.Open() ??
                                   archive.CreateEntry(DocumentRelsInfo.Path).Open())
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

            using (Stream stream = archive.GetEntry("word/document.xml")?.Open() ??
                                   archive.CreateEntry("word/document.xml").Open())
            {
                stream.SetLength(0);
                document.Save(stream);
            }

            XElement packageRelation = archive.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(
                    T + "Override",
                    new XAttribute("PartName", "/word/footer2.xml"),
                    new XAttribute("ContentType", FooterContentType)));

            using (Stream stream = archive.GetEntry(ContentTypesInfo.Path)?.Open() ??
                                   archive.CreateEntry(ContentTypesInfo.Path).Open())
            {
                stream.SetLength(0);
                packageRelation.Save(stream);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="publisher">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static XDocument Footer1([NotNull] string publisher)
        {
            if (publisher is null)
                throw new ArgumentNullException(nameof(publisher));

            return
                new XDocument(
                    Declaration,
                    new XElement(
                        W + "ftr",
                        new XElement(
                            W + "p",
                            new XElement(
                                W + "pPr",
                                new XElement(
                                    W + "jc",
                                    new XAttribute(W + "val", "right"))),
                            new XElement(
                                W + "r",
                                new XElement(
                                    W + "t",
                                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                                    new XText($"{publisher} | "))),
                            new XElement(
                                W + "fldSimple",
                                new XAttribute(W + "instr", " PAGE   \\* MERGEFORMAT "),
                                new XElement(W + "r",
                                    new XElement(W + "t",
                                        new XText(string.Empty)))))));
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="website">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static XDocument Footer2([NotNull] string website)
        {
            if (website is null)
                throw new ArgumentNullException(nameof(website));

            return
                new XDocument(
                    Declaration,
                    new XElement(
                        W + "ftr",
                        new XElement(
                            W + "p",
                            new XElement(
                                W + "pPr",
                                new XElement(
                                    W + "jc",
                                    new XAttribute(W + "val", "left"))),
                            new XElement(
                                W + "fldSimple",
                                new XAttribute(W + "instr", " PAGE   \\* MERGEFORMAT "),
                                new XElement(W + "r",
                                    new XElement(W + "t",
                                        new XText(string.Empty)))),
                            new XElement(
                                W + "r",
                                new XElement(
                                    W + "t",
                                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                                    new XText($" | {website}"))))));
        }
    }
}