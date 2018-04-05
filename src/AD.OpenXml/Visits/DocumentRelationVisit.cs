using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Structures;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <inheritdoc />
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public sealed class DocumentRelationVisit : IOpenXmlPackageVisit
    {
        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <inheritdoc />
        public OpenXmlPackageVisitor Result { get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public DocumentRelationVisit(OpenXmlPackageVisitor subject, int documentRelationId)
        {
            Document document = Execute(subject.Document, documentRelationId);

            Result = subject.With(document);
        }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="documentRelationId"></param>
        /// <returns>
        /// The updated document node of the source file.
        /// </returns>
        [Pure]
        private static Document Execute(Document document, int documentRelationId)
        {
            XElement modifiedDocument =
                new XElement(
                    document.Content.Name,
                    document.Content.Attributes().Select(x => Update(x, documentRelationId)),
                    document.Content.Elements().Select(x => Update(x, documentRelationId)));

            ChartInfo[] chartMapping =
                document.Charts
                        .Select(x => x.WithOffset(documentRelationId))
                        .ToArray();

            ImageInfo[] imageMapping =
                document.Images
                        .Select(x => x.WithOffset(documentRelationId))
                        .ToArray();

            HyperlinkInfo[] hyperlinkMapping =
                document.Hyperlinks
                        .Select(x => x.WithOffset(documentRelationId))
                        .ToArray();

            return new Document(modifiedDocument, chartMapping, imageMapping, hyperlinkMapping);
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