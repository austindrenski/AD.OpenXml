using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// Marshals charts from the 'chart#.xml' files of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public class OpenXmlChartVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// Active version of 'word/document.xml'.
        /// </summary>
        public override XElement Document { get; }

        /// <summary>
        /// Active version of 'word/_rels/document.xml.rels'.
        /// </summary>
        public override XElement DocumentRelations { get; }

        /// <summary>
        /// Active version of '[Content_Types].xml'.
        /// </summary>
        public override XElement ContentTypes { get; }

        /// <summary>
        /// Active version of word/charts/chart#.xml.
        /// </summary>
        public override IEnumerable<ChartInformation> Charts { get; }

        /// <summary>
        /// Returns the last document relation identifier in use by the container.
        /// </summary>
        public override int DocumentRelationId { get; }

        /// <summary>
        /// Marshal footnotes from the source document into the container.
        /// </summary>
        /// <returns>The updated document node of the source file.</returns>
        public OpenXmlChartVisitor(OpenXmlVisitor subject, int documentRelationId) : base(subject)
        {
            (Document, DocumentRelations, ContentTypes, Charts, DocumentRelationId)
                = Execute(subject.File, subject.Document, subject.DocumentRelations, subject.ContentTypes, subject.Charts, documentRelationId);
        }

        /// <summary>
        /// Marshal footnotes from the source document into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <param name="document">The document node of the source file containing any modifications made to this point.</param>
        /// <param name="documentRelations"></param>
        /// <param name="contentTypes"></param>
        /// <param name="charts"></param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        public static (XElement Document, XElement DocumentRelations, XElement ContentTypes, IEnumerable<ChartInformation> Charts, int DocumentRelationId)
            Execute(DocxFilePath file, XElement document, XElement documentRelations, XElement contentTypes, IEnumerable<ChartInformation> charts, int documentRelationId)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }
            if (contentTypes is null)
            {
                throw new ArgumentNullException(nameof(contentTypes));
            }
            if (charts is null)
            {
                throw new ArgumentNullException(nameof(charts));
            }

            var chartMapping =
                documentRelations.Descendants(P + "Relationship")
                                 .Where(x => x.Attribute("Target")?.Value.StartsWith("charts/") ?? false)
                                 .Select(
                                     x => new
                                     {
                                         SourceId = x.Attribute("Id")?.Value,
                                         SourceIdNumeric = x.Attribute("Id")?.Value.ParseInt().GetValueOrDefault() ?? 0,
                                         SourceName = x.Attribute("Target")?.Value
                                     })
                                 .Select(
                                     x => new
                                     {
                                         x.SourceId,
                                         x.SourceName,
                                         ResultId = $"rId{documentRelationId + x.SourceIdNumeric}",
                                         ResultName = $"charts/chart{documentRelationId + x.SourceIdNumeric}.xml",
                                         NewNumericId = documentRelationId + x.SourceIdNumeric
                                     })
                                 .Select(
                                     x => new
                                     {
                                         x,
                                         ChartInformation = new ChartInformation(x.ResultName, file.ReadAsXml($"word/{x.SourceName}"))
                                     })
                                 .Select(
                                     x =>
                                     {
                                         x.ChartInformation.Chart.Descendants(C + "externalData").Remove();
                                         return x;
                                     })
                                 .ToArray();

            XElement modifiedSourceContent = document.Clone();

            foreach (var map in chartMapping)
            {
                modifiedSourceContent = modifiedSourceContent.ChangeXAttributeValues(C + "chart", R + "id", map.x.SourceId, map.x.ResultId);
            }

            XElement modifiedContentTypes =
                new XElement(
                    contentTypes.Name,
                    contentTypes.Attributes(),
                    contentTypes.Elements(),
                    chartMapping.Select(x =>
                        new XElement(T + "Override",
                            new XAttribute("PartName", $"/word/{x.ChartInformation.Name}"),
                            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))));

            XElement modifiedDocumentRelations =
                new XElement(
                    documentRelations.Name,
                    documentRelations.Attributes(),
                    documentRelations.Elements()
                                     .Where(x => x.Attribute("Type")?.Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink") ?? false),
                    chartMapping.Select(x =>
                        new XElement(P + "Relationship",
                            new XAttribute("Id", x.x.ResultId),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                            new XAttribute("Target", $"{x.ChartInformation.Name}"))));

            IEnumerable<ChartInformation> modifiedCharts = 
                chartMapping.Select(x => x.ChartInformation)
                            .ToImmutableArray();

            int updatedDocumentRelationId =
                chartMapping.Any()
                    ? chartMapping.Max(x => x.x.NewNumericId)
                    : documentRelationId;

            return (modifiedSourceContent, modifiedDocumentRelations, modifiedContentTypes, modifiedCharts, updatedDocumentRelationId);
        }
    }
}