using System;
using System.Collections.Generic;
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
        [NotNull]
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull]
        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        [NotNull]
        private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Active version of 'word/document.xml'.
        /// </summary>
        [NotNull]
        private XElement _document;

        /// <summary>
        /// Active version of '[Content_Types].xml'.
        /// </summary>
        [NotNull]
        private XElement _contentTypes;

        /// <summary>
        /// Active version of 'word/footnotes.xml'.
        /// </summary>
        [NotNull]
        private XElement _footnotes;

        /// <summary>
        /// Active version of 'word/_rels/document.xml.rels'.
        /// </summary>
        [NotNull]
        private XElement _documentRelations;

        /// <summary>
        /// Active version of word/charts/chart#.xml.
        /// </summary>
        [NotNull]
        [ItemNotNull]
        private List<Tuple<string, XElement>> _charts;

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
        /// <param name="result">The file to which changes can be saved.</param>
        public OpenXmlContainer([NotNull] DocxFilePath result)
        {
            _document = result.ReadAsXml("word/document.xml");

            _contentTypes = result.ReadAsXml("[Content_Types].xml");

            _footnotes = result.ReadAsXml("word/footnotes.xml");

            _documentRelations = result.ReadAsXml("word/_rels/document.xml.rels");

            _charts =
                _documentRelations.Elements()
                                  .Select(x => x.Attribute("Target")?.Value)
                                  .Where(x => x?.StartsWith("charts/") ?? false)
                                  .Select(x => Tuple.Create(x, result.ReadAsXml($"word/{x}")))
                                  .ToList();
        }

        private OpenXmlContainer(
            [NotNull] XElement document,
            [NotNull] XElement contentTypes,
            [NotNull] XElement footnotes,
            [NotNull] XElement documentRelations,
            [NotNull][ItemNotNull] IEnumerable<Tuple<string, XElement>> charts
            )
        {
            _document = document.Clone();

            _contentTypes = contentTypes.Clone();

            _footnotes = footnotes.Clone();

            _documentRelations = documentRelations.Clone();

            _charts = charts.Select(x => Tuple.Create(x.Item1, x.Item2.Clone())).ToList();
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
            foreach (Tuple<string, XElement> chart in _charts)
            {
                chart.Item2.WriteInto(resultPath, $"word/{chart.Item1}");
            }
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="files">The files from which content is copied.</param>
        public void MergeDocuments([NotNull][ItemNotNull] IEnumerable<DocxFilePath> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            foreach (DocxFilePath file in files)
            {
                MergeDocuments(file);
            }
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        public void MergeDocuments([NotNull] DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            XElement source = file.ReadAsXml("word/document.xml");

            source = MarshalContentFrom(source);

            source = MergeFootnotesFrom(file, source);

            source = MergeChartsFrom(file, source);

            IEnumerable<XElement> sourceContent =
                source.Elements()
                      .Single()
                      .Elements();

            _document.Elements()
                     .First()
                     .Add(sourceContent);
        }

        /// <summary>
        /// Marshal content from the source document to be added into the container.
        /// </summary>
        /// <param name="sourceContent">The document node of the source file containing any modifications made to this point.</param>
        /// <returns>The updated document node of the source file.</returns>
        [NotNull]
        [MustUseReturnValue]
        private XElement MarshalContentFrom([NotNull] XElement sourceContent)
        {
            if (sourceContent is null)
            {
                throw new ArgumentNullException(nameof(sourceContent));
            }

            XElement source =
                sourceContent.RemoveRsidAttributes()
                             .RemoveRunPropertiesFromParagraphProperties()
                             .RemoveByAll(W + "proofErr")
                             .RemoveByAll(W + "bookmarkStart")
                             .RemoveByAll(W + "bookmarkEnd")
                             .MergeRuns()
                             .ChangeBoldToStrong()
                             .ChangeItalicToEmphasis()
                             .ChangeUnderlineToTableCaption()
                             .ChangeUnderlineToFigureCaption()
                             .ChangeUnderlineToSourceNote()
                             .ChangeSuperscriptToReference()
                             .HighlightInsertRequests()
                             .AddLineBreakToHeadings()
                             .SetTableStyles()
                             .RemoveByAll(W + "rFonts")
                             .RemoveByAll(W + "sz")
                             .RemoveByAll(W + "szCs")
                             .RemoveByAll(W + "u")
                             .RemoveByAllIfEmpty(W + "rPr")
                             .RemoveByAllIfEmpty(W + "pPr")
                             .RemoveByAllIfEmpty(W + "t")
                             .RemoveByAllIfEmpty(W + "r")
                             .RemoveByAllIfEmpty(W + "p");

            source.Descendants(W + "rPr")
                  .Where(
                      x =>
                          x.Elements(W + "rStyle")
                           .Attributes(W + "val")
                           .Any(y => y.Value.Equals("FootnoteReference")))
                  .SelectMany(
                      x =>
                          x.Descendants()
                           .Where(y => !y.Attribute(W + "val")?.Value.Equals("FootnoteReference") ?? false))
                  .Remove();

            source.Descendants(W + "p").Attributes().Remove();
            source.Descendants(W + "tr").Attributes().Remove();
            source.Descendants(W + "hideMark").Remove();
            source.Descendants(W + "noWrap").Remove();
            source.Descendants(W + "pPr").Where(x => !x.HasElements).Remove();
            source.Descendants(W + "rPr").Where(x => !x.HasElements).Remove();
            source.Descendants(W + "spacing").Remove();

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
        /// <returns>The updated document node of the source file.</returns>
        [NotNull]
        [MustUseReturnValue]
        private XElement MergeFootnotesFrom([NotNull] DocxFilePath file, [NotNull] XElement sourceContent)
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
                return sourceContent;
            }

            IEnumerable<XElement> sourceFootnotes =
                file.ReadAsXml("word/footnotes.xml")
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
        /// <param name="sourceContent">The document node of the source file containing any modifications made to this point.</param>
        /// <returns>The updated document node of the source file.</returns>
        [NotNull]
        [MustUseReturnValue]
        private static (XElement, XElement, XElement, List<Tuple<string, XElement>>) MergeChartsFrom([NotNull] DocxFilePath source, [NotNull] XElement sourceContent, int currentDocumentRelationId, XElement documentRelations, XElement contentTypes, List<Tuple<string, XElement>> charts)
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
                    chartMapping.Select(
                        x =>
                            new XElement(T + "Override",
                                new XAttribute("PartName", $"/word/{x.ResultName}"),
                                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));

            foreach (var map in chartMapping)
            {
                sourceContent = sourceContent.ChangeXAttributeValues(C + "chart", R + "id", map.SourceId, map.ResultId);

                contentTypes.Add(
                    new XElement(T + "Override",
                        new XAttribute("PartName", $"/word/{map.ResultName}"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));

                documentRelations.Add(
                    new XElement(P + "Relationship",
                        new XAttribute("Id", map.ResultId),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                        new XAttribute("Target", $"{map.ResultName}")));

                XElement chart = source.ReadAsXml($"word/{map.SourceName}");
                chart.Descendants(C + "externalData").Remove();
                _charts.Add(Tuple.Create(map.ResultName, chart));
            }

            return sourceContent;
        }
    }
}