using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.OpenXml.Visitors;
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
            (var document, var documentRelations, var contentTypes, var charts, var images) =
                Execute(
                    subject.Document.RemoveRsidAttributes(),
                    subject.DocumentRelations,
                    subject.ContentTypes,
                    subject.Charts,
                    subject.Images,
                    documentRelationId);

            Result =
                new OpenXmlPackageVisitor(
                    contentTypes,
                    document,
                    documentRelations,
                    subject.Footnotes,
                    subject.FootnoteRelations,
                    subject.Styles,
                    subject.Numbering,
                    subject.Theme1,
                    charts,
                    images);
        }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="documentRelations"></param>
        /// <param name="contentTypes"></param>
        /// <param name="charts"></param>
        /// <param name="images"></param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        private static (XElement Document, XElement DocumentRelations, XElement ContentTypes, ChartInformation[] Charts, ImageInformation[] Images) Execute(XElement document, XElement documentRelations, XElement contentTypes, IEnumerable<ChartInformation> charts, IEnumerable<ImageInformation> images, uint documentRelationId)
        {
            XElement modifiedDocumentRelations =
                new XElement(
                    documentRelations.Name,
                    documentRelations.Attributes(),
                    documentRelations.Elements().Select(x => UpdateElement(x, documentRelationId)));

            (string oldId, string newId)[] documentRelationMapping =
                modifiedDocumentRelations.Elements()
                                         .Attributes("Id")
                                         .Select(x => (oldId: $"rId{int.Parse(x.Value.Substring(3)) - documentRelationId}", newId: x.Value))
                                         .ToArray();

            foreach (XAttribute item in document.Descendants().Attributes(R + "id"))
            {
                (string _, string newId) = documentRelationMapping.Single(x => x.oldId == (string) item);

                item.SetValue(newId);

                item.Parent?
                    .Ancestors(W + "drawing")
                    .Descendants()
                    .Attributes("id")
                    .SingleOrDefault()?
                    .SetValue(newId);
            }

            foreach (XAttribute item in document.Descendants().Attributes(R + "embed"))
            {
                (string _, string newId) = documentRelationMapping.Single(x => x.oldId == (string) item);

                item.SetValue(newId);

                item.Parent?
                    .Ancestors(W + "drawing")
                    .Descendants()
                    .Attributes("embed")
                    .SingleOrDefault()?
                    .SetValue(newId);
            }

            ChartInformation[] chartMapping =
                charts.Select(x => x.WithOffset(documentRelationId))
                      .ToArray();

            ImageInformation[] imageMapping =
                images.Select(x => x.WithOffset(documentRelationId))
                      .ToArray();

            XElement modifiedContentTypes =
                new XElement(
                    contentTypes.Name,
                    contentTypes.Attributes(),
                    new XElement(
                        T + "Default",
                        new XAttribute("Extension", "jpeg"),
                        new XAttribute("ContentType", "image/jpeg")),
                    new XElement(
                        T + "Default",
                        new XAttribute("Extension", "png"),
                        new XAttribute("ContentType", "image/png")),
                    new XElement(
                        T + "Default",
                        new XAttribute("Extension", "svg"),
                        new XAttribute("ContentType", "image/svg")),
                    chartMapping.Select(x => x.ContentTypeEntry));

            return (document, modifiedDocumentRelations, modifiedContentTypes, chartMapping, imageMapping);
        }

        private static XElement UpdateElement(XElement e, uint offset)
        {
            uint candidate = offset + uint.Parse(e.Attribute("Id").Value.Substring(3));

            return
                new XElement(
                    e.Name,
                    e.Attributes().Select(x => UpdateAttribute(x, candidate)),
                    e.Elements().Select(x => UpdateElement(x, offset)));
        }

        private static XAttribute UpdateAttribute(XAttribute a, uint candidate)
        {
            switch (a.Name.LocalName)
            {
                case "Id":
                {
                    return new XAttribute(a.Name, $"rId{candidate}");
                }
                case "Target" when TargetChart.IsMatch(a.Value):
                {
                    return new XAttribute(a.Name, $"charts/chart{candidate}.xml");
                }
                case "Target" when TargetImage.IsMatch(a.Value):
                {
                    Match m = TargetImage.Match(a.Value);
                    return new XAttribute(a.Name, $"media/image{candidate}.{m.Groups["extension"].Value}");
                }
                default:
                {
                    return a;
                }
            }
        }
    }
}