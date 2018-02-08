using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using AD.IO;
using AD.IO.Paths;
using AD.IO.Streams;
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
        public static async Task<MemoryStream> AddFooters([NotNull] this Task<MemoryStream> stream)
        {
            return await AddFooters(await stream);
        }

        /// <summary>
        /// Add footers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static async Task<MemoryStream> AddFooters([NotNull] this MemoryStream stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            MemoryStream result = await stream.CopyPure();

            // Modify [Content_Types].xml
            XElement packageRelation = result.ReadAsXml("[Content_Types].xml");

            packageRelation.Descendants(T + "Override")
                           .Where(x => x.Attribute("PartName")?.Value.StartsWith("/word/footer") ?? false)
                           .Remove();

            result = await packageRelation.WriteInto(result, "[Content_Types].xml");

            // Modify document.xml.rels and grab the current header id number
            XElement documentRelation = result.ReadAsXml("word/_rels/document.xml.rels");

            documentRelation.Descendants(P + "Relationship")
                            .Where(x => x.Attribute("Target")?.Value.Contains("footer") ?? false)
                            .Remove();

            int currentFooterId =
                documentRelation.Elements().Count();

            result = await documentRelation.WriteInto(result, "word/_rels/document.xml.rels");

            // Modify document.xml
            XElement document = result.ReadAsXml();

            document.Descendants(W + "sectPr")
                    .Elements(W + "footerReference")
                    .Remove();

            result = await document.WriteInto(result, "word/document.xml");

            // Add footers
            result = await AddEvenPageFooter(result, $"rId{++currentFooterId}");
            result = await AddOddPageFooter(result, $"rId{++currentFooterId}");

            return result;
        }

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

            int currentFooterId =
                documentRelation.Elements().Count();

            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            // Modify document.xml
            XElement document = toFilePath.ReadAsXml();

            document.Descendants(W + "sectPr")
                    .Elements(W + "footerReference")
                    .Remove();

            document.WriteInto(toFilePath, "word/document.xml");

            // Add footers
            toFilePath.AddEvenPageFooter($"rId{++currentFooterId}");
            toFilePath.AddOddPageFooter($"rId{++currentFooterId}");
        }

        [Pure]
        [NotNull]
        private static async Task<MemoryStream> AddOddPageFooter([NotNull] MemoryStream stream, [NotNull] string footerId)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (footerId is null)
            {
                throw new ArgumentNullException(nameof(footerId));
            }

            MemoryStream result = await stream.CopyPure();

            result = await Footer1.WriteInto(result, "word/footer1.xml");

            XElement documentRelation = result.ReadAsXml("word/_rels/document.xml.rels");

            documentRelation.Add(
                new XElement(P + "Relationship",
                             new XAttribute("Id", footerId),
                             new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                             new XAttribute("Target", "footer1.xml")));

            result = await documentRelation.WriteInto(result, "word/_rels/document.xml.rels");

            XElement document = result.ReadAsXml();

            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.Elements(W + "headerReference")
                                 .Last()
                                 .AddAfterSelf(
                                     new XElement(W + "footerReference",
                                                  new XAttribute(W + "type", "default"),
                                                  new XAttribute(R + "id", footerId)));
            }

            result = await document.WriteInto(result, "word/document.xml");

            XElement packageRelation = result.ReadAsXml("[Content_Types].xml");

            packageRelation.Add(
                new XElement(T + "Override",
                             new XAttribute("PartName", "/word/footer1.xml"),
                             new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));

            result = await packageRelation.WriteInto(result, "[Content_Types].xml");

            return result;
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

            Footer1.WriteInto(toFilePath, "word/footer1.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(P + "Relationship",
                             new XAttribute("Id", footerId),
                             new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                             new XAttribute("Target", "footer1.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml();
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.Elements(W + "headerReference")
                                 .Last()
                                 .AddAfterSelf(
                                     new XElement(W + "footerReference",
                                                  new XAttribute(W + "type", "default"),
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

        [Pure]
        [NotNull]
        private static async Task<MemoryStream> AddEvenPageFooter([NotNull] MemoryStream stream, [NotNull] string footerId)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (footerId is null)
            {
                throw new ArgumentNullException(nameof(footerId));
            }

            MemoryStream result = await stream.CopyPure();

            result = await Footer2.WriteInto(result, "word/footer2.xml");

            XElement documentRelation = result.ReadAsXml("word/_rels/document.xml.rels");

            documentRelation.Add(
                new XElement(P + "Relationship",
                             new XAttribute("Id", footerId),
                             new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                             new XAttribute("Target", "footer2.xml")));

            result = await documentRelation.WriteInto(result, "word/_rels/document.xml.rels");

            XElement document = result.ReadAsXml();

            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.Elements(W + "headerReference")
                                 .Last()
                                 .AddAfterSelf(new XElement(W + "footerReference",
                                                            new XAttribute(W + "type", "even"),
                                                            new XAttribute(R + "id", footerId)));
            }

            result = await document.WriteInto(result, "word/document.xml");

            XElement packageRelation = result.ReadAsXml("[Content_Types].xml");

            packageRelation.Add(
                new XElement(T + "Override",
                             new XAttribute("PartName", "/word/footer2.xml"),
                             new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")));

            result = await packageRelation.WriteInto(result, "[Content_Types].xml");

            return result;
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

            Footer2.WriteInto(toFilePath, "word/footer2.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");

            documentRelation.Add(
                new XElement(P + "Relationship",
                             new XAttribute("Id", footerId),
                             new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"),
                             new XAttribute("Target", "footer2.xml")));

            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml();

            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.Elements(W + "headerReference")
                                 .Last()
                                 .AddAfterSelf(new XElement(W + "footerReference",
                                                            new XAttribute(W + "type", "even"),
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