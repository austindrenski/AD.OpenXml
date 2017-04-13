using System;
using System.Collections.Generic;
using System.Collections.Immutable;
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
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Active version of 'word/document.xml'.
        /// </summary>
        [NotNull]
        private readonly XElement _document;

        /// <summary>
        /// Active version of 'word/_rels/document.xml.rels'.
        /// </summary>
        [NotNull]
        private readonly XElement _documentRelations;

        /// <summary>
        /// Active version of '[Content_Types].xml'.
        /// </summary>
        [NotNull]
        private readonly XElement _contentTypes;

        /// <summary>
        /// Active version of 'word/footnotes.xml'.
        /// </summary>
        [NotNull]
        private readonly XElement _footnotes;
        
        /// <summary>
        /// Active version of word/charts/chart#.xml.
        /// </summary>
        [NotNull]
        private readonly IImmutableList<(string Name, XElement Chart)> _charts;

        /// <summary>
        /// Returns the last footnote identifier currently in use by the container.
        /// </summary>
        private readonly int _currentFootnoteId;

        /// <summary>
        /// Initializes an <see cref="OpenXmlContainer"/> by reading document parts into memory.
        /// </summary>
        /// <param name="result">The file to which changes can be saved.</param>
        public OpenXmlContainer([NotNull] DocxFilePath result)
        {
            _document = result.ReadAsXml("word/document.xml");
            _documentRelations = result.ReadAsXml("word/_rels/document.xml.rels");
            _contentTypes = result.ReadAsXml("[Content_Types].xml");
            _footnotes = result.ReadAsXml("word/footnotes.xml");
            _charts = result.ReadAsXml("word/_rels/document.xml.rels")
                            .Elements()
                            .Select(x => x.Attribute("Target")?.Value)
                            .Where(x => x?.StartsWith("charts/") ?? false)
                            .Select(x => (Name: x, Chart: result.ReadAsXml($"word/{x}")))
                            .ToImmutableList();

            _currentFootnoteId =
                _footnotes.Elements(W + "footnote")
                          .Attributes(W + "id")
                          .Select(x => int.Parse(x.Value))
                          .Max();
        }

        /// <summary>
        /// Initializes an <see cref="OpenXmlContainer"/> by reading document parts into memory.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="documentRelations"></param>
        /// <param name="contentTypes"></param>
        /// <param name="footnotes"></param>
        /// <param name="charts"></param>
        /// <param name="currentFootnoteId"></param>
        public OpenXmlContainer([NotNull] XElement document, XElement documentRelations, XElement contentTypes, XElement footnotes, IEnumerable<(string Name, XElement Chart)> charts, int currentFootnoteId)
        {
            _document = document.Clone();
            _documentRelations = documentRelations.Clone();
            _contentTypes = contentTypes.Clone();
            _footnotes = footnotes.Clone();
            _charts = charts.Select(x => (Name: x.Name, Chart: x.Chart.Clone())).ToImmutableArray();
        }

        /// <summary>
        /// Saves any modifications to <paramref name="resultPath"/>. This operation will overwrite any existing content for the modified parts.
        /// </summary>
        /// <param name="resultPath">The path to which modified parts are written.</param>
        public void Save([NotNull] DocxFilePath resultPath)
        {
            _document.WriteInto(resultPath, "word/document.xml");
            _footnotes.WriteInto(resultPath, "word/footnotes.xml");
            _contentTypes.WriteInto(resultPath, "[Content_Types].xml");
            _documentRelations.WriteInto(resultPath, "word/_rels/document.xml.rels");
            foreach ((string Name, XElement Chart) chart in _charts)
            {
                chart.Chart.WriteInto(resultPath, $"word/{chart.Name}");
            }
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="files">The files from which content is copied.</param>
        [Pure]
        [NotNull]
        public OpenXmlContainer MergeDocuments([NotNull][ItemNotNull] IEnumerable<DocxFilePath> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            return files.Aggregate(this, (current, next) => current.MergeDocuments(next));
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        [Pure]
        [NotNull]
        public OpenXmlContainer MergeDocuments([NotNull] DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }
           
            XElement modifiedSourceContent = 
                MarshalContentFrom(file);

            (XElement footnoteModifiedSourceContent, XElement modifiedSourceFootnotes, int updatedFootnoteId) = 
                MergeFootnotesFrom(file, modifiedSourceContent, _currentFootnoteId);

            (XElement chartAndFootnoteModifiedSourceContent, XElement modifiedDocumentRelations, XElement modifiedContentTypes, IEnumerable<(string Name, XElement Chart)> modifiedCharts) =
                MergeChartsFrom(file, footnoteModifiedSourceContent, _contentTypes, _documentRelations, _charts);

            XElement mergedContent =
                new XElement(
                    _document.Name,
                    _document.Attributes(),
                    new XElement(W + "body",
                        _document.Element(W + "body")?.Elements(),
                        chartAndFootnoteModifiedSourceContent.Element(W + "body")?.Elements()));

            XElement mergedFootnotes =
                new XElement(
                    _footnotes.Name,
                    _footnotes.Attributes(),
                    _footnotes.Elements(),
                    modifiedSourceFootnotes?.Elements());

            return new OpenXmlContainer(
                mergedContent,
                modifiedDocumentRelations,
                modifiedContentTypes,
                mergedFootnotes,
                modifiedCharts,
                updatedFootnoteId);

        }

        /// <summary>
        /// Marshal content from the source document to be added into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        [NotNull]
        private static XElement MarshalContentFrom([NotNull] DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            XElement source =
                file.ReadAsXml("word/document.xml")
                    .RemoveRsidAttributes()
                    .RemoveRunPropertiesFromParagraphProperties()
                    .RemoveByAll(W + "proofErr")
                    .RemoveByAll(W + "bookmarkStart")
                    .RemoveByAll(W + "bookmarkEnd")
                    .RemoveByAll(W + "tblPrEx")
                    .RemoveByAll(W + "spacing")
                    .RemoveByAll(W + "lang")
                    .RemoveByAll(W + "numPr")
                    .RemoveByAll(W + "hideMark")
                    .RemoveByAll(W + "noWrap")
                    .RemoveByAll(W + "rFonts")
                    .RemoveByAll(W + "sz")
                    .RemoveByAll(W + "szCs")
                    .RemoveByAll(W + "color")
                    .RemoveByAll(W + "lastRenderedPageBreak")
                    .RemoveByAll(x => x.Name.Equals(W + "br") && (x.Attribute(W + "type")?.Value.Equals("page", StringComparison.OrdinalIgnoreCase) ?? false))
                    .RemoveByAll(x => x.Name.Equals(W + "pStyle") && (x.Attribute(W + "val")?.Value.Equals("BodyTextSSFinal", StringComparison.OrdinalIgnoreCase) ?? false))
                    .MergeRuns()
                    .ChangeBoldToStrong()
                    .ChangeItalicToEmphasis()
                    .ChangeUnderlineToTableCaption()
                    .ChangeUnderlineToFigureCaption()
                    .ChangeUnderlineToSourceNote()
                    .ChangeSuperscriptToReference()
                    .HighlightInsertRequests()
                    //.AddLineBreakToHeadings()
                    .SetTableStyles()
                    .RemoveByAll(W + "u")
                    .RemoveByAllIfEmpty(W + "tcPr")
                    .RemoveByAllIfEmpty(W + "rPr")
                    .RemoveByAllIfEmpty(W + "pPr")
                    .RemoveByAllIfEmpty(W + "t")
                    .RemoveByAllIfEmpty(W + "r")
                    .RemoveByAll(x => x.Name.Equals(W + "p") && !x.HasElements && (!x.Parent?.Name.Equals(W + "tc") ?? false));

            foreach (XElement paragraphProperties in source.Descendants(W + "pPr").Where(x => x.Elements(W + "pStyle").Count() > 1))
            {
                IEnumerable<XElement> styles = paragraphProperties.Elements(W + "pStyle").ToArray();
                styles.Remove();
                paragraphProperties.AddFirst(styles.Distinct());
            }

            foreach (XElement runProperties in source.Descendants(W + "rPr").Where(x => x.Elements(W + "rStyle").Count() > 1))
            {
                IEnumerable<XElement> styles = runProperties.Elements(W + "rStyle").ToArray();
                styles.Remove();
                IEnumerable<XElement> distinct = styles.Distinct().ToArray();
                if (distinct.Any(x => x.Attribute(W + "val")?.Value.Equals("FootnoteReference") ?? false))
                {
                    distinct = distinct.Where(x => x.Attribute(W + "val")?.Value.Equals("FootnoteReference") ?? false);
                }
                runProperties.AddFirst(distinct);
            }

            source.Descendants(W + "p").Attributes().Remove();

            source.Descendants(W + "tr").Attributes().Remove();

            if (source.Element(W + "body")?.Elements().First().Name == W + "sectPr")
            {
                source.Element(W + "body")?.Elements().First().Remove();
            }

            source.Descendants(W + "hyperlink").Remove();

            return source;
        }

        /// <summary>
        /// Merge footnotes from the source document into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <param name="sourceContent">The document node of the source file containing any modifications made to this point.</param>
        /// <param name="currentFootnoteId">The last footnote number currently in use by the container.</param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        private static (XElement SourceContent, XElement SourceFootnotes, int UpdatedFootnoteId) MergeFootnotesFrom([NotNull] DocxFilePath file, [NotNull] XElement sourceContent, int currentFootnoteId)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            if (sourceContent is null)
            {
                throw new ArgumentNullException(nameof(sourceContent));
            }

            // TODO: Make the return type of ReadAsXml() a nullable singleton.
            try
            {
                file.ReadAsXml("word/footnotes.xml");
            }
            catch
            {
                return (SourceContent: sourceContent, SourceFootnotes: null, UpdatedFootnoteId: currentFootnoteId);
            }

            XElement sourceFootnotes =
                file.ReadAsXml("word/footnotes.xml")
                    .RemoveRsidAttributes();

            sourceFootnotes.Descendants(W + "p")
                           .Attributes()
                           .Remove();

            sourceFootnotes.Descendants(W + "hyperlink")
                           .Remove();

            var footnoteMapping =
                sourceFootnotes.Elements(W + "footnote")
                               .Attributes(W + "id")
                               .Select(x => x.Value)
                               .Select(int.Parse)
                               .Where(x => x > 0)
                               .OrderByDescending(x => x)
                               .Select(
                                   x => new
                                   {
                                       oldId = $"{x}",
                                       newId = $"{currentFootnoteId + x}",
                                       newNumericId = currentFootnoteId + x
                                   })
                               .ToArray();

            foreach (var map in footnoteMapping)
            {
                sourceContent =
                    sourceContent.ChangeXAttributeValues(W + "footnoteReference", W + "Id", map.oldId, map.newId);

                sourceFootnotes =
                    sourceFootnotes.ChangeXAttributeValues(W + "footnote", W + "id", map.oldId, map.newId);
            }

            return (SourceContent: sourceContent, SourceFootnotes: sourceFootnotes, UpdatedFootnoteId: footnoteMapping.Max(x => x.newNumericId));
        }

        /// <summary>
        /// Merge footnotes from the source document into the container.
        /// </summary>
        /// <param name="source">The file from which content is copied.</param>
        /// <param name="sourceContent">The document node of the source file containing any modifications made to this point.</param>
        /// <param name="contentTypes"></param>
        /// <param name="documentRelations"></param>
        /// <param name="charts"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        private static (XElement SourceContent, XElement DocumentRelations, XElement ContentTypes, IEnumerable<(string Name, XElement Chart)> Charts) 
            MergeChartsFrom([NotNull] DocxFilePath source, XElement sourceContent, XElement contentTypes, XElement documentRelations, IEnumerable<(string Name, XElement Chart)> charts)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
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
                source.ReadAsXml("word/_rels/document.xml.rels")
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
                                            Chart = source.ReadAsXml($"word/{x.SourceName}")
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