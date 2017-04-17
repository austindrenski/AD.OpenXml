using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
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
        /// Active version of 'word/_rels/footnotes.xml.rels'.
        /// </summary>
        [NotNull]
        private readonly XElement _footnoteRelations;

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
        /// Returns the last footnote hyperlink identifier currently in use by the container.
        /// </summary>
        private readonly int _currentFootnoteRelationId;

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
                MarshalFootnotesFrom(file, modifiedSourceContent, _currentFootnoteId);


            (XElement chartAndFootnoteModifiedSourceContent, XElement modifiedDocumentRelations, XElement modifiedContentTypes, IEnumerable<(string Name, XElement Chart)> modifiedCharts) =
                MarshalChartsFrom(file, footnoteModifiedSourceContent, _contentTypes, _documentRelations, _charts);

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
                    _footnotes.Elements()
                              .Union(
                                  modifiedSourceFootnotes?.Elements() ?? Enumerable.Empty<XElement>(),
                                  XNode.EqualityComparer));

            XElement mergedFootnoteRelations = new XElement("a");

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
        /// <remarks>This method delegates its work to the <see cref="MarshalContentFromExtensions.MarshalContentFrom(DocxFilePath)"/> extension method.</remarks>
        [Pure]
        [NotNull]
        private static XElement MarshalContentFrom([NotNull] DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            return file.MarshalContentFrom();
        }

        /// <summary>
        /// Marshal footnotes from the source document into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <param name="sourceContent">The document node of the source file containing any modifications made to this point.</param>
        /// <param name="currentFootnoteId">The last footnote number currently in use by the container.</param>
        /// <returns>The updated document node of the source file.</returns>
        /// <remarks>This method delegates its work to the <see cref="MarshalFootnotesFromExtensions.MarshalFootnotesFrom(DocxFilePath, XElement, int)"/> extension method.</remarks>
        [Pure]
        private static (XElement SourceContent, XElement SourceFootnotes, int UpdatedFootnoteId) 
            MarshalFootnotesFrom([NotNull] DocxFilePath file, [NotNull] XElement sourceContent, int currentFootnoteId)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            if (sourceContent is null)
            {
                throw new ArgumentNullException(nameof(sourceContent));
            }

            return file.MarshalFootnotesFrom(sourceContent, currentFootnoteId);
        }

        /// <summary>
        /// Marshal footnotes from the source document into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <param name="sourceContent">The document node of the source file containing any modifications made to this point.</param>
        /// <param name="contentTypes"></param>
        /// <param name="documentRelations"></param>
        /// <param name="charts"></param>
        /// <returns>The updated document node of the source file.</returns>
        /// <remarks>This method delegates its work to the MarshalChartsFromExtensions.MarshalChartsFrom(...) extension method.</remarks>
        [Pure]
        private static (XElement SourceContent, XElement DocumentRelations, XElement ContentTypes, IEnumerable<(string Name, XElement Chart)> Charts)
            MarshalChartsFrom([NotNull] DocxFilePath file, XElement sourceContent, XElement contentTypes, XElement documentRelations, IEnumerable<(string Name, XElement Chart)> charts)
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

            return file.MarshalChartsFrom(sourceContent, contentTypes, documentRelations, charts);
        }
    }
}