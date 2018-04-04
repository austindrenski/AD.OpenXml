using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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

        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] private static readonly Regex TargetChart = new Regex("charts/chart(?<id>[0-9]+)\\.xml$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        [NotNull] private static readonly Regex TargetImage = new Regex("media/image(?<id>[0-9]+)\\.(?<extension>jpeg|png|svg)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <inheritdoc />
        ///  <summary>
        ///  </summary>
        public OpenXmlPackageVisitor Result { get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public DocumentRelationVisit(OpenXmlPackageVisitor subject, uint documentRelationId)
        {
            (var document, var charts, var images, var hyperlinks) =
                Execute(
                    subject.Document.Content,
                    subject.Document.Charts,
                    subject.Document.Images,
                    subject.Document.Hyperlinks,
                    documentRelationId);

            Document resultDoc = new Document(document, charts, images, hyperlinks);

            Result = subject.With(document: resultDoc);
        }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="charts"></param>
        /// <param name="images"></param>
        /// <param name="hyperlinks"></param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        private static (XElement Document, ChartInfo[] Charts, ImageInfo[] Images, HyperlinkInfo[] Hyperlinks)
            Execute(
                XElement document,
                IEnumerable<ChartInfo> charts,
                IEnumerable<ImageInfo> images,
                IEnumerable<HyperlinkInfo> hyperlinks,
                uint documentRelationId)
        {
            XElement modifiedDocument =
                new XElement(
                    document.Name,
                    document.Attributes().Select(x => Update(x, documentRelationId)),
                    document.Elements().Select(x => Update(x, documentRelationId)));

            ChartInfo[] chartMapping =
                charts.Select(x => x.WithOffset(documentRelationId))
                      .ToArray();

            ImageInfo[] imageMapping =
                images.Select(x => x.WithOffset(documentRelationId))
                      .ToArray();

            HyperlinkInfo[] hyperlinkMapping =
                hyperlinks.Select(x => x.WithOffset(documentRelationId))
                          .ToArray();

            return (modifiedDocument, chartMapping, imageMapping, hyperlinkMapping);
        }

        private static XObject Update(XObject node, uint offset)
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
                    return new XAttribute(a.Name, $"rId{offset + uint.Parse(a.Value.Substring(3))}");
                }
                default:
                {
                    return node;
                }
            }
        }
    }
}