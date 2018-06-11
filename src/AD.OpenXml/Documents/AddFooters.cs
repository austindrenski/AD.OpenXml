using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
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
        [NotNull] private static readonly string MimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

        /// <summary>
        /// The schema type for an OpenXML footer relationship.
        /// </summary>
        [NotNull] private static readonly string RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";

        /// <summary>
        /// Add footers to a Word document.
        /// </summary>
        [Pure]
        [NotNull]
        public static Package AddFooters([NotNull] this Package package, [NotNull] string publisher, [NotNull] string website)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            if (publisher is null)
                throw new ArgumentNullException(nameof(publisher));

            if (website is null)
                throw new ArgumentNullException(nameof(website));

            foreach (PackagePart part in package.GetParts())
            {
                if (part.ContentType == MimeType)
                    package.DeletePart(part.Uri);

                if (part.ContentType != Document.ContentType)
                    continue;

                foreach (PackageRelationship relationship in part.GetRelationshipsByType(RelationshipType))
                {
                    part.DeleteRelationship(relationship.Id);
                }
            }

            using (Stream stream = package.GetPart(Document.PartName).GetStream())
            {
                XElement document = XElement.Load(stream);
                document.Descendants(W + "footerReference").Remove();
                stream.SetLength(0);
                document.Save(stream);
            }

            // Add footers
            AddEvenPageFooter(package, website);
            AddOddPageFooter(package, publisher);

            return package;
        }

        private static void AddOddPageFooter([NotNull] Package package, [NotNull] string publisher)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            if (publisher is null)
                throw new ArgumentNullException(nameof(publisher));

            Uri headerUri = new Uri("/word/footer1.xml", UriKind.Relative);
            PackagePart headerPart = package.CreatePart(headerUri, MimeType);

            using (Stream stream = headerPart.GetStream())
            {
                Footer1(publisher).Save(stream);
            }

            PackageRelationship relationship =
                package.GetPart(Document.PartName).CreateRelationship(headerUri, TargetMode.Internal, RelationshipType);

            using (Stream stream = package.GetPart(Document.PartName).GetStream())
            {
                XElement document = XElement.Load(stream);

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
                                                 new XAttribute(R + "id", relationship.Id)));
                    }
                    else
                    {
                        sectionProperties.Add(
                            new XElement(
                                W + "footerReference",
                                new XAttribute(W + "type", "default"),
                                new XAttribute(R + "id", relationship.Id)));
                    }
                }

                stream.SetLength(0);

                document.Save(stream);
            }
        }

        private static void AddEvenPageFooter([NotNull] Package package, [NotNull] string website)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            if (website is null)
                throw new ArgumentNullException(nameof(website));

            Uri headerUri = new Uri("/word/footer2.xml", UriKind.Relative);
            PackagePart headerPart = package.CreatePart(headerUri, MimeType);

            using (Stream stream = headerPart.GetStream())
            {
                Footer2(website).Save(stream);
            }

            PackageRelationship relationship =
                package.GetPart(Document.PartName).CreateRelationship(headerUri, TargetMode.Internal, RelationshipType);

            using (Stream stream = package.GetPart(Document.PartName).GetStream())
            {
                XElement document = XElement.Load(stream);

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
                                                 new XAttribute(R + "id", relationship.Id)));
                    }
                    else
                    {
                        sectionProperties.Add(
                            new XElement(
                                W + "footerReference",
                                new XAttribute(W + "type", "even"),
                                new XAttribute(R + "id", relationship.Id)));
                    }
                }

                stream.SetLength(0);

                document.Save(stream);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="publisher"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        private static XDocument Footer1([NotNull] string publisher)
            => new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(
                    W + "ftr",
                    new XAttribute(XNamespace.Xmlns + "w", W),
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

        /// <summary>
        ///
        /// </summary>
        /// <param name="website"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        private static XDocument Footer2([NotNull] string website)
            => new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(
                    W + "ftr",
                    new XAttribute(XNamespace.Xmlns + "w", W),
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