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
        public static async Task<MemoryStream> AddFooters([NotNull] this Task<MemoryStream> stream, [NotNull] string publisher, [NotNull] string website)
        {
            return await AddFooters(await stream, publisher, website);
        }

        /// <summary>
        /// Add footers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static async Task<MemoryStream> AddFooters([NotNull] this MemoryStream stream, [NotNull] string publisher, [NotNull] string website)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            MemoryStream result = await stream.CopyPure();

            // Remove footers from [Content_Types].xml
            result =
                await result.ReadXml(ContentTypesInfo.Path)
                            .Recurse(x => (string) x.Attribute(ContentTypesInfo.Attributes.ContentType) != FooterContentType)
                            .WriteIntoAsync(result, ContentTypesInfo.Path);

            // Remove footers from document.xml.rels
            result =
                await result.ReadXml(DocumentRelsInfo.Path)
                            .Recurse(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Type) != FooterRelationshipType)
                            .WriteIntoAsync(result, DocumentRelsInfo.Path);

            // Remove footers from document.xml
            result =
                await result.ReadXml()
                            .Recurse(x => x.Name != W + "footerReference")
                            .WriteIntoAsync(result, "word/document.xml");

            // Store the current relationship id number
            int currentRelationshipId =
                result.ReadXml(DocumentRelsInfo.Path)
                      .Elements()
                      .Attributes("Id")
                      .Select(x => int.Parse(x.Value.Substring(3)))
                      .DefaultIfEmpty(0)
                      .Max();

            // Add footers
            result = await AddEvenPageFooter(result, $"rId{++currentRelationshipId}", website);
            result = await AddOddPageFooter(result, $"rId{++currentRelationshipId}", publisher);

            return result;
        }

        [Pure]
        [NotNull]
        private static async Task<MemoryStream> AddOddPageFooter([NotNull] MemoryStream stream, [NotNull] string footerId, [NotNull] string publisher)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (footerId is null)
            {
                throw new ArgumentNullException(nameof(footerId));
            }

            if (publisher is null)
            {
                throw new ArgumentNullException(nameof(publisher));
            }

            MemoryStream result = await stream.CopyPure();

            XElement footer1 = Footer1.Clone();
            footer1.Element(W + "p").Elements(W + "r").First().Element(W + "t").Value = publisher;

            result = await footer1.WriteIntoAsync(result, "word/footer1.xml");

            XElement documentRelation = result.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(P + "Relationship",
                             new XAttribute("Id", footerId),
                             new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                             new XAttribute("Target", "footer1.xml")));

            result = await documentRelation.WriteIntoAsync(result, DocumentRelsInfo.Path);

            XElement document = result.ReadXml();

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

            result = await document.WriteIntoAsync(result, "word/document.xml");

            XElement packageRelation = result.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(T + "Override",
                             new XAttribute("PartName", "/word/footer1.xml"),
                             new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));

            result = await packageRelation.WriteIntoAsync(result, ContentTypesInfo.Path);

            return result;
        }

        [Pure]
        [NotNull]
        private static async Task<MemoryStream> AddEvenPageFooter([NotNull] MemoryStream stream, [NotNull] string footerId, [NotNull] string website)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (footerId is null)
            {
                throw new ArgumentNullException(nameof(footerId));
            }

            if (website is null)
            {
                throw new ArgumentNullException(nameof(website));
            }

            MemoryStream result = await stream.CopyPure();

            XElement footer2 = Footer2.Clone();
            footer2.Element(W + "p").Elements(W + "r").Last().Element(W + "t").Value = website;

            result = await footer2.WriteIntoAsync(result, "word/footer2.xml");

            XElement documentRelation = result.ReadXml(DocumentRelsInfo.Path);

            documentRelation.Add(
                new XElement(P + "Relationship",
                             new XAttribute("Id", footerId),
                             new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                             new XAttribute("Target", "footer2.xml")));

            result = await documentRelation.WriteIntoAsync(result, DocumentRelsInfo.Path);

            XElement document = result.ReadXml();

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

            result = await document.WriteIntoAsync(result, "word/document.xml");

            XElement packageRelation = result.ReadXml(ContentTypesInfo.Path);

            packageRelation.Add(
                new XElement(T + "Override",
                             new XAttribute("PartName", "/word/footer2.xml"),
                             new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));

            result = await packageRelation.WriteIntoAsync(result, ContentTypesInfo.Path);

            return result;
        }
    }
}