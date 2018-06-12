using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
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
        [NotNull] private static readonly string ContentType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";

        /// <summary>
        /// The schema type for an OpenXML header relationship.
        /// </summary>
        [NotNull] private static readonly string RelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";

        /// <summary>
        /// Add headers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static Package AddHeaders([NotNull] this Package package, [NotNull] string title)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            if (title is null)
                throw new ArgumentNullException(nameof(title));

            Package result = package.ToPackage(FileAccess.ReadWrite);

            foreach (PackagePart part in result.GetParts())
            {
                if (part.ContentType == ContentType)
                    result.DeletePart(part.Uri);

                if (part.ContentType != Document.ContentType)
                    continue;

                foreach (PackageRelationship relationship in part.GetRelationshipsByType(RelationshipType))
                {
                    part.DeleteRelationship(relationship.Id);
                }
            }

            using (Stream stream = result.GetPart(Document.PartUri).GetStream())
            {
                XElement document = XElement.Load(stream);
                document.Descendants(W + "headerReference").Remove();
                stream.SetLength(0);
                document.Save(stream);
            }

            AddOddPageHeader(result);
            AddEvenPageHeader(result, title);

            return result;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="package"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private static void AddOddPageHeader([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            Uri headerUri = new Uri("/word/header2.xml", UriKind.Relative);
            PackagePart headerPart = package.CreatePart(headerUri, ContentType);

            using (Stream stream = headerPart.GetStream())
            {
                Header2().Save(stream);
            }

            PackageRelationship relationship =
                package.GetPart(Document.PartUri)
                       .CreateRelationship(headerUri, TargetMode.Internal, RelationshipType);

            using (Stream stream = package.GetPart(Document.PartUri).GetStream())
            {
                XElement document = XElement.Load(stream);

                foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
                {
                    sectionProperties.AddFirst(
                        new XElement(
                            W + "headerReference",
                            new XAttribute(W + "type", "default"),
                            new XAttribute(R + "id", relationship.Id)));
                }

                stream.SetLength(0);

                document.Save(stream);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="package"></param>
        /// <param name="title"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private static void AddEvenPageHeader([NotNull] Package package, [NotNull] string title)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            if (title is null)
                throw new ArgumentNullException(nameof(title));

            Uri headerUri = new Uri("/word/header1.xml", UriKind.Relative);
            PackagePart headerPart = package.CreatePart(headerUri, ContentType);

            using (Stream stream = headerPart.GetStream())
            {
                Header1(title).Save(stream);
            }

            PackageRelationship relationship =
                package.GetPart(Document.PartUri)
                       .CreateRelationship(headerUri, TargetMode.Internal, RelationshipType);

            using (Stream stream = package.GetPart(Document.PartUri).GetStream())
            {
                XElement document = XElement.Load(stream);

                foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
                {
                    sectionProperties.AddFirst(
                        new XElement(
                            W + "headerReference",
                            new XAttribute(W + "type", "even"),
                            new XAttribute(R + "id", relationship.Id)));
                }

                stream.SetLength(0);

                document.Save(stream);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="title"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        private static XDocument Header1([NotNull] string title)
            => new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(
                    W + "hdr",
                    new XAttribute(XNamespace.Xmlns + "w", W),
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

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        private static XDocument Header2()
            => new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(
                    W + "hdr",
                    new XAttribute(XNamespace.Xmlns + "w", W),
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