using System;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Structures;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public static class DocumentRelationVisit
    {
        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        [NotNull]
        public static OpenXmlPackageVisitor VisitDocRels([NotNull] this OpenXmlPackageVisitor subject, int documentRelationId)
        {
            if (subject is null)
                throw new ArgumentNullException(nameof(subject));

            Document document = subject.Document;

            XElement modifiedDocument =
                new XElement(
                    document.Content.Name,
                    document.Content.Attributes().Select(x => Update(x, documentRelationId)),
                    document.Content.Elements().Select(x => Update(x, documentRelationId)));

            return
                subject.With(
                    new Document(
                        modifiedDocument,
                        document.Charts.Select(x => x.WithOffset(documentRelationId)),
                        document.Images.Select(x => x.WithOffset(documentRelationId)),
                        document.Hyperlinks.Select(x => x.WithOffset(documentRelationId))));
        }

        private static XObject Update(XObject node, int offset)
        {
            switch (node)
            {
                case XElement e:
                {
                    return
                        new XElement(
                            e.Name,
                            e.Attributes().Select(x => Update(x, offset)),
                            e.Elements().Select(x => Update(x, offset)),
                            e.HasElements ? null : e.Value);
                }
                case XAttribute a when a.Name == R + "id" || a.Name == R + "embed":
                {
                    return new XAttribute(a.Name, $"rId{offset + int.Parse(a.Value.Substring(3))}");
                }
                default:
                {
                    return node;
                }
            }
        }
    }
}