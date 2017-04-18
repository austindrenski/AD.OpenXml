using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
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
        /// Returns the last document relation identifier in use by the container.
        /// </summary>
        private readonly int _currentDocumentRelationId;

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
            _document = 
                result.ReadAsXml() ?? throw new FileNotFoundException("document.xml");

            _contentTypes = 
                result.ReadAsXml("[Content_Types].xml") ?? throw new FileNotFoundException("[Content_Types].xml");

            _footnotes =
                result.ReadAsXml("word/footnotes.xml") ?? new XElement(W + "footnotes");

            _documentRelations = 
                result.ReadAsXml("word/_rels/document.xml.rels") ?? new XElement(P + "Relationships");

            _footnoteRelations = 
                result.ReadAsXml("word/_rels/footnotes.xml.rels") ?? new XElement(P + "Relationships");

            _charts =
                _documentRelations.Elements()
                                  .Select(x => x.Attribute("Target")?.Value)
                                  .Where(x => x?.StartsWith("charts/") ?? false)
                                  .Select(x => (Name: x, Chart: result.ReadAsXml($"word/{x}")))
                                  .ToImmutableList();

            _currentDocumentRelationId =
                _documentRelations.Elements(P + "Relationship")
                                  .Attributes("Id")
                                  .Select(x => x.Value.ParseInt() ?? 0)
                                  .DefaultIfEmpty(0)
                                  .Max();

            _currentFootnoteId =
                _footnotes.Elements(W + "footnote")
                          .Attributes(W + "id")
                          .Select(x => x.Value.ParseInt() ?? 0)
                          .DefaultIfEmpty(0)
                          .Max();

            _currentFootnoteRelationId =
                _footnoteRelations.Elements(P + "Relationship")
                                  .Attributes("Id")
                                  .Select(x => x.Value.ParseInt() ?? 0)
                                  .DefaultIfEmpty(0)
                                  .Max();
        }

        /// <summary>
        /// Initializes an <see cref="OpenXmlContainer"/> by reading document parts into memory.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="documentRelations"></param>
        /// <param name="contentTypes"></param>
        /// <param name="footnotes"></param>
        /// <param name="foonoteRelations"></param>
        /// <param name="charts"></param>
        /// <param name="currentFootnoteId"></param>
        /// <param name="currentFootnoteRelationId"></param>
        /// <param name="currentDocumentRelationId"></param>
        public OpenXmlContainer([NotNull] XElement document, XElement documentRelations, XElement contentTypes, XElement footnotes, XElement foonoteRelations, IEnumerable<(string Name, XElement Chart)> charts, int currentFootnoteId, int currentFootnoteRelationId, int currentDocumentRelationId)
        {
            _document = document.Clone();
            _documentRelations = documentRelations.Clone();
            _contentTypes = contentTypes.Clone();
            _footnotes = footnotes.Clone();
            _footnoteRelations = foonoteRelations.Clone();
            _charts = charts.Select(x => (Name: x.Name, Chart: x.Chart.Clone())).ToImmutableArray();
            // TODO: Why isn't this needed? Actually, why does this proactively break the document?
            //_currentFootnoteId = currentFootnoteId;
            _currentFootnoteRelationId = currentFootnoteRelationId;
            _currentDocumentRelationId = currentDocumentRelationId;
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
            _footnoteRelations.WriteInto(resultPath, "word/_rels/footnotes.xml.rels");
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

            XElement sourceContent1 =
                file.MarshalContentFrom();

            (XElement sourceContent2, XElement sourceFootnotes1, int updatedFootnoteId) =
                file.MarshalFootnotesFrom(sourceContent1, _currentFootnoteId);

            (XElement sourceFootnotes2, XElement footnoteRelations1, int updatedFootnoteRelationId) = 
                file.MarshalFootnoteHyperlinksFrom(sourceFootnotes1, _currentFootnoteRelationId);

            (XElement sourceContent3, XElement documentRelations1, int documentRelationId1) =
                file.MarshalContentHyperlinksFrom(sourceContent2, _currentDocumentRelationId);

            (XElement sourceContent4, XElement documentRelations2, XElement contentTypes1, IEnumerable<(string Name, XElement Chart)> charts1, int documentRelationId2) =
                file.MarshalChartsFrom(sourceContent3, _contentTypes, _charts, documentRelationId1);

            XElement resultContent =
                new XElement(
                    _document.Name,
                    _document.Attributes(),
                    new XElement(W + "body",
                        _document.Element(W + "body")?.Elements(),
                        sourceContent4.Element(W + "body")?.Elements()));

            XElement resultFootnotes =
                new XElement(
                    _footnotes.Name,
                    _footnotes.Attributes(),
                    _footnotes.Elements()
                              .Union(
                                  sourceFootnotes2?.Elements() ?? Enumerable.Empty<XElement>(),
                                  XNode.EqualityComparer));

            XElement resultFootnoteRelations =
                new XElement(
                    _footnoteRelations.Name,
                    _footnoteRelations.Attributes(),
                    _footnoteRelations.Elements()
                                      .Union(
                                          footnoteRelations1?.Elements() ?? Enumerable.Empty<XElement>(),
                                          XNode.EqualityComparer));

            XElement resultDocumentRelations =
                new XElement(
                    _documentRelations.Name,
                    _documentRelations.Attributes(),
                    _documentRelations.Elements()
                                      .Union(
                                          documentRelations1?.Elements() ?? Enumerable.Empty<XElement>(),
                                          XNode.EqualityComparer)
                                      .Union(
                                          documentRelations2?.Elements() ?? Enumerable.Empty<XElement>(),
                                          XNode.EqualityComparer));

            return new OpenXmlContainer(
                resultContent,
                resultDocumentRelations,
                contentTypes1,
                resultFootnotes,
                resultFootnoteRelations,
                charts1,
                updatedFootnoteId,
                updatedFootnoteRelationId,
                documentRelationId2);
        }
    }
}