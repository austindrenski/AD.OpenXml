using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// This class serves as a container to encapsulate XML components of a Word document.
    /// </summary>
    [PublicAPI]
    public class OpenXmlContainer
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// word/document.xml
        /// </summary>
        private readonly XElement _sourceDocument;
        private XElement _document;

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        private readonly XElement _sourceContentTypes;
        private XElement _contentTypes;
        
        /// <summary>
        /// word/footnotes.xml
        /// </summary>
        private readonly XElement _sourceFootnotes;
        private XElement _footnotes;
        
        /// <summary>
        /// word/_rels/document.xml.rels
        /// </summary>
        private readonly XElement _sourceDocumentRelations;
        private XElement _documentRelations;
        
        /// <summary>
        /// word/charts/chart#.xml
        /// </summary>
        private readonly IImmutableList<XElement> _sourceCharts;
        private XElement[] _charts;

        /// <summary>
        /// word/charts/_rels/chart#.xml.rels
        /// </summary>
        private readonly IImmutableList<XElement> _sourceChartRelations;
        private XElement[] _chartRelations;

        /// <summary>
        /// Returns the numeric value of the highest footnote identifier of the form 'w:id="#"'.
        /// </summary>
        private int CurrentFootnoteId
        {
            get
            {
                return _footnotes.Elements(W + "footnote")
                                 .Attributes(W + "id")
                                 .Select(x => int.Parse(x.Value))
                                 .Max();
            }
        }

        /// <summary>
        /// Returns the numeric component of the highest document relation identifier of the form 'Id="rId#"'.
        /// </summary>
        private int CurrentDocumentRelationId
        {
            get
            {
                return _documentRelations.Elements()
                                         .Attributes("Id")
                                         .Select(x => x.Value.ParseInt().GetValueOrDefault())
                                         .Max();
            }
        }

        /// <summary>
        /// Initializes an <see cref="OpenXmlContainer"/> by reading document parts into memory.
        /// </summary>
        /// <param name="result"></param>
        public OpenXmlContainer(DocxFilePath result)
        {
            _sourceDocument = result.ReadAsXml("word/document.xml");
            _document = _sourceDocument.Clone();

            _sourceContentTypes = result.ReadAsXml("[Content_Types].xml");
            _contentTypes = _sourceContentTypes.Clone();

            _sourceFootnotes = result.ReadAsXml("word/footnotes.xml");
            _footnotes = _sourceFootnotes.Clone();

            _sourceDocumentRelations = result.ReadAsXml("word/_rels/document.xml.rels");
            _documentRelations = _sourceDocumentRelations.Clone();
            
            _sourceCharts =
                _documentRelations.Elements()
                                  .Select(x => x.Attribute("Target")?.Value)
                                  .Where(x => x?.StartsWith("charts/") ?? false)
                                  .Select(x => result.ReadAsXml($"word/{x}"))
                                  .ToImmutableArray();
            _charts = _sourceCharts.Clone().ToArray();

            _sourceChartRelations =
                _documentRelations.Elements()
                                  .Select(x => x.Attribute("Target")?.Value)
                                  .Where(x => x?.StartsWith("charts/") ?? false)
                                  .Select(x => result.ReadAsXml($"word/charts/_rels/{x}.rels"))
                                  .ToImmutableArray();
            _chartRelations = _sourceChartRelations.Clone().ToArray();
        }

        /// <summary>
        /// Saves any modifications to <paramref name="resultPath"/>. This operation will overwrite any existing content for the modified parts.
        /// </summary>
        /// <param name="resultPath">The path to which modified parts are written.</param>
        public void Save([NotNull] DocxFilePath resultPath)
        {
            if (!XNode.DeepEquals(_sourceDocument, _document))
            {
                _document.WriteInto(resultPath, "word/document.xml");
            }
            if (!XNode.DeepEquals(_sourceFootnotes, _footnotes))
            {
                _footnotes.WriteInto(resultPath, "word/footnotes.xml");
            }
            if (!XNode.DeepEquals(_sourceContentTypes, _contentTypes))
            {
                _contentTypes.WriteInto(resultPath, "[Content_Types].xml");
            }
            if (!XNode.DeepEquals(_documentRelations, _documentRelations))
            {
                _documentRelations.WriteInto(resultPath, "word/_rels/document.xml.rels");
            }
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="source">The file from which content is copied.</param>
        /// <param name="result"></param> // TODO: get rid of this parameter in favor of a transform that isn't disk-dependent.
        public void MergeDocuments([NotNull] DocxFilePath source, [NotNull] DocxFilePath result)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            XElement sourceDocument =
                source.ReadAsXml("word/document.xml")
                      .Process508From();

            sourceDocument = MergeFootnotesFrom(source, sourceDocument);

            sourceDocument = MergeChartsFrom(source, result, sourceDocument);

            IEnumerable<XElement> sourceContent =
                sourceDocument.Elements()
                              .Single()
                              .Elements();

            _document.Elements()
                     .First()
                     .Add(sourceContent);
        }

        /// <summary>
        /// Merge footnotes from the source document into the container.
        /// </summary>
        /// <param name="source">The file from which content is copied.</param>
        /// <param name="sourceContent">The document node of the source file.</param>
        private XElement MergeFootnotesFrom([NotNull] DocxFilePath source, [NotNull] XElement sourceContent)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            if (sourceContent is null)
            {
                throw new ArgumentNullException(nameof(sourceContent));
            }

            // TODO: Make the return type of ReadAsXml() a nulalble singleton.
            try
            {
                source.ReadAsXml("word/footnotes.xml");
            }
            catch
            {
                return sourceContent;
            }

            IEnumerable<XElement> sourceFootnotes =
                source.ReadAsXml("word/footnotes.xml")
                      .Elements(W + "footnote")
                      .RemoveRsidAttributes()
                      .ToArray();

            sourceFootnotes.Descendants(W + "p")
                           .Attributes()
                           .Remove();

            sourceFootnotes.Descendants(W + "hyperlink")
                           .Remove();

            var footnoteMapping =
                sourceFootnotes.Attributes(W + "id")
                               .Select(x => x.Value)
                               .Select(int.Parse)
                               .Where(x => x > 0)
                               .OrderByDescending(x => x)
                               .Select(
                                   x => new
                                   {
                                       oldId = $"{x}",
                                       newId = $"{CurrentFootnoteId + x}"
                                   });

            foreach (var map in footnoteMapping)
            {
                sourceContent =
                    sourceContent.ChangeXAttributeValues(W + "footnoteReference", W + "Id", map.oldId, map.newId);

                sourceFootnotes =
                    sourceFootnotes.ChangeXAttributeValues(W + "footnote", W + "id", map.oldId, map.newId)
                                   .ToArray();
            }

            _footnotes.Add(sourceFootnotes);

            return sourceContent;
        }

        /// <summary>
        /// Merge footnotes from the source document into the container.
        /// </summary>
        /// <param name="source">The file from which content is copied.</param>
        /// <param name="result">The into which content is copied.</param>
        /// <param name="sourceContent"></param>
        private XElement MergeChartsFrom([NotNull] DocxFilePath source, DocxFilePath result, XElement sourceContent)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            if (sourceContent is null)
            {
                throw new ArgumentNullException(nameof(sourceContent));
            }

            var chartMapping =
                source.ReadAsXml("word/_rels/document.xml.rels")
                      .Descendants(P + "Relationship")
                      .Where(x => x.Attribute("Target")?.Value.StartsWith("charts/") ?? false)
                      .Select(
                          x => new
                          {
                              SourceIdNumeric = x.Attribute("Id")?.Value.ParseInt().GetValueOrDefault() ?? 0,
                              SourceId = x.Attribute("Id")?.Value,
                              ResultId = $"rId{CurrentDocumentRelationId + x.Attribute("Id")?.Value.ParseInt().GetValueOrDefault()}",
                              SourceName = x.Attribute("Target")?.Value.Replace("charts/", null)
                          })
                      .ToArray();

            foreach (var map in chartMapping)
            {
                sourceContent = sourceContent.ChangeXAttributeValues(C + "chart", R + "id", map.SourceId, map.ResultId);
                TransferChart(source, result, map.SourceId, map.ResultId, map.SourceName, chartMapping.Max(x => x.SourceIdNumeric));
            }

            return sourceContent;
        }

        internal static void TransferChart(DocxFilePath source, DocxFilePath result, string fromId, string toId, string inputName, int currentMaxId)
        {
            int nextChartNumber;
            using (ZipArchive archive = ZipFile.OpenRead(result))
            {
                nextChartNumber =
                    archive.Entries
                           .Where(x => x.FullName.StartsWith("word/charts/chart"))
                           .Select(x => x.FullName.Replace("word/charts/chart", null).Replace(".xml", null))
                           .Select(int.Parse)
                           .DefaultIfEmpty(0)
                           .Max() + 1;
            }

            XElement contentTypes = result.ReadAsXml("[Content_Types].xml");
            contentTypes.Add(
                new XElement(T + "Override",
                    new XAttribute("PartName", $"/word/charts/chart{nextChartNumber}.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));
            contentTypes.WriteInto(result, "[Content_Types].xml");

            XElement documentRelation = result.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", toId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                    new XAttribute("Target", $"charts/chart{nextChartNumber}.xml")));
            documentRelation.WriteInto(result, "word/_rels/document.xml.rels");

            XElement chart =
                source.ReadAsXml($"word/charts/{inputName}");
            chart.Descendants(C + "externalData").Remove();
            chart.WriteInto(result, $"word/charts/chart{nextChartNumber}.xml");
        }
    }
}