using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
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
        [NotNull] private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] private static readonly XNamespace WP = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

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
        public DocumentRelationVisit(OpenXmlPackageVisitor subject, int documentRelationId)
        {
            (var document, var documentRelations, var contentTypes, var charts, var images) =
                Execute(
                    subject.Document,
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
        private static (XElement Document, XElement DocumentRelations, XElement ContentTypes, IEnumerable<ChartInformation> Charts, IEnumerable<ImageInformation> Images) Execute(XElement document, XElement documentRelations, XElement contentTypes, IEnumerable<ChartInformation> charts, IEnumerable<ImageInformation> images, int documentRelationId)
        {
            var documentRelationMapping =
                documentRelations.Descendants(P + "Relationship")
                                 .Where(x => (string) x.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" ||
                                             (string) x.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ||
                                             (string) x.Attribute("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
                                 .Select(
                                     x => new
                                     {
                                         Id = x.Attribute("Id"),
                                         Type = x.Attribute("Type"),
                                         Target = x.Attribute("Target"),
                                         TargetMode = x.Attribute("TargetMode")
                                     })
                                 .OrderBy(x => x.Id.Value.ParseInt())
                                 .Select(
                                     (x, i) => new
                                     {
                                         oldId = x.Id,
                                         newId = new XAttribute("Id", $"rId{documentRelationId + i}"),
                                         x.Type,
                                         oldTarget = x.Target,
                                         newTarget =
                                             x.Target.Value.StartsWith("charts/")
                                                 ? new XAttribute("Target", $"charts/chart{documentRelationId + i}.xml")
                                                 : x.Target.Value.StartsWith("media/")
                                                     ? new XAttribute("Target", $"media/image{documentRelationId + i}{x.Target.Value.Substring(x.Target.Value.LastIndexOf('.'))}")
                                                     : x.Target,
                                         x.TargetMode
                                     })
                                 .ToArray();

            ChartInformation[] chartMapping =
                documentRelationMapping
                    .Where(x => x.oldTarget.Value.StartsWith("charts/"))
                    .Select(
                        x => new
                        {
                            x.newId,
                            x.oldTarget,
                            x.newTarget
                        })
                    .OrderBy(x => x.newId.Value.ParseInt())
                    .Select(
                        // TODO: fix this...should be single
                        x => new ChartInformation(x.newTarget.Value, charts.First(y => y.Name == x.oldTarget.Value).Chart))
                    .Select(
                        x =>
                        {
                            x.Chart.Descendants(C + "externalData").Remove();
                            return x;
                        })
                    .ToArray();

            ImageInformation[] imageMapping =
                documentRelationMapping
                    .Where(x => x.oldTarget.Value.StartsWith("media/"))
                    .Select(
                        x => new
                        {
                            x.newId,
                            x.oldTarget,
                            x.newTarget
                        })
                    .OrderBy(x => x.newId.Value.ParseInt())
                    .Select(
                        // TODO: fix this...should be single
                        x => new ImageInformation(x.newTarget.Value, images.First(y => y.Name == x.oldTarget.Value).Image))
                    .ToArray();

            XElement modifiedDocument =
                document.RemoveRsidAttributes();

            foreach (XAttribute item in modifiedDocument.Descendants().Attributes(R + "id"))
            {
                var map = documentRelationMapping.SingleOrDefault(x => (string) x.oldId == (string) item);

                if (map is null)
                {
                    continue;
                }

                item.SetValue((string) map.newId);

                item.Parent?
                    .Ancestors(W + "drawing")
                    .Descendants()
                    .Attributes("id")
                    // TODO: fix this...should be single
                    .FirstOrDefault()?
                    .SetValue(map.newId.Value.ParseInt().ToString());
            }

            foreach (XAttribute item in modifiedDocument.Descendants().Attributes(R + "embed"))
            {
                var map = documentRelationMapping.SingleOrDefault(x => (string) x.oldId == (string) item);

                if (map is null)
                {
                    continue;
                }

                item.SetValue((string) map.newId);

                item.Parent?
                    .Ancestors(W + "drawing")
                    .Descendants()
                    .Attributes("embed")
                    // TODO: fix this...should be single
                    .FirstOrDefault()?
                    .SetValue(map.newId.Value.ParseInt().ToString());
            }

            XElement modifiedDocumentRelations =
                new XElement(
                    documentRelations.Name,
                    documentRelationMapping.Select(
                        x =>
                            new XElement(
                                P + "Relationship",
                                x.newId,
                                x.Type,
                                x.newTarget,
                                x.TargetMode)));

            XElement modifiedContentTypes =
                new XElement(
                    contentTypes.Name,
                    new XElement(T + "Default",
                        new XAttribute("Extension", "png"),
                        new XAttribute("ContentType", "image/png")),
                    new XElement(T + "Default",
                        new XAttribute("Extension", "jpeg"),
                        new XAttribute("ContentType", "image/jpeg")),
                    chartMapping.Select(
                        x =>
                            new XElement(T + "Override",
                                new XAttribute("PartName", $"/word/{x.Name}"),
                                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))));

            // TODO: does modified content types need to know about image defaults?

            return (modifiedDocument, modifiedDocumentRelations, modifiedContentTypes, chartMapping, imageMapping);
        }
    }
}