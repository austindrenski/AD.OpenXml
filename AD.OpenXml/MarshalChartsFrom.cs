using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Marshals charts from the 'chart#.xml' files of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public static class MarshalChartsFromExtensions
    {
        /// <summary>
        /// Represents the 'c:' prefix seen in the markup for chart[#].xml
        /// </summary>
        [NotNull]
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of [Content_Types].xml
        /// </summary>
        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of document.xml.
        /// </summary>
        [NotNull]
        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// The namespace declared on the [Content_Types].xml
        /// </summary>
        [NotNull]
        private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        /// <summary>
        /// Marshal footnotes from the source document into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <param name="sourceContent">The document node of the source file containing any modifications made to this point.</param>
        /// <param name="contentTypes"></param>
        /// <param name="documentRelations"></param>
        /// <param name="charts"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        public static (XElement SourceContent, XElement DocumentRelations, XElement ContentTypes, IEnumerable<(string Name, XElement Chart)> Charts)
            MarshalChartsFrom([NotNull] this DocxFilePath file, XElement sourceContent, XElement contentTypes, XElement documentRelations, IEnumerable<(string Name, XElement Chart)> charts)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            if (sourceContent is null)
            {
                throw new ArgumentNullException(nameof(sourceContent));
            }
            if (contentTypes is null)
            {
                throw new ArgumentNullException(nameof(contentTypes));
            }
            if (documentRelations is null)
            {
                throw new ArgumentNullException(nameof(documentRelations));
            }
            if (charts is null)
            {
                throw new ArgumentNullException(nameof(charts));
            }

            int currentDocumentRelationId =
                documentRelations.Elements(P + "Relationship")
                                 .Attributes("Id")
                                 .Select(x => x.Value.ParseInt().GetValueOrDefault())
                                 .Max();

            var chartMapping =
                file.ReadAsXml("word/_rels/document.xml.rels")
                    .Descendants(P + "Relationship")
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
                            ResultId = $"rId{currentDocumentRelationId + x.SourceIdNumeric}",
                            ResultName = $"charts/chart{currentDocumentRelationId + x.SourceIdNumeric}.xml"
                        })
                    .ToArray();

            XElement modifiedContentTypes =
                new XElement(
                    contentTypes.Name,
                    contentTypes.Attributes(),
                    contentTypes.Elements(),
                    chartMapping.Select(x =>
                        new XElement(T + "Override",
                            new XAttribute("PartName", $"/word/{x.ResultName}"),
                            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))));

            XElement modifiedDocumentRelations =
                new XElement(
                    documentRelations.Name,
                    documentRelations.Attributes(),
                    documentRelations.Elements(),
                    chartMapping.Select(x =>
                        new XElement(P + "Relationship",
                            new XAttribute("Id", x.ResultId),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                            new XAttribute("Target", $"{x.ResultName}"))));

            IEnumerable<(string Name, XElement Chart)> modifiedCharts =
                charts
                    .Concat(
                        chartMapping.Select(
                                        x => new
                                        {
                                            Name = x.ResultName,
                                            Chart = file.ReadAsXml($"word/{x.SourceName}")
                                        })
                                    .Select(
                                        x =>
                                        {
                                            x.Chart.Descendants(C + "externalData").Remove();
                                            return (Name: x.Name, Chart: x.Chart);
                                        }));

            XElement modifiedSourceContent = sourceContent.Clone();

            foreach (var map in chartMapping)
            {
                modifiedSourceContent = modifiedSourceContent.ChangeXAttributeValues(C + "chart", R + "id", map.SourceId, map.ResultId);
            }

            return (SourceContent: modifiedSourceContent, DocumentRelations: modifiedDocumentRelations, ContentTypes: modifiedContentTypes, Charts: modifiedCharts);
        }
    }
}