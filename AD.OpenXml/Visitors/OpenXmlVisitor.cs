using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// This class serves as a container to encapsulate XML components of a Word document.
    /// </summary>
    [PublicAPI]
    public class OpenXmlVisitor
    {
        /// <summary>
        /// Represents the 'c:' prefix seen in the markup for chart[#].xml
        /// </summary>
        [NotNull]
        protected static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of [Content_Types].xml
        /// </summary>
        [NotNull]
        protected static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of document.xml.
        /// </summary>
        [NotNull]
        protected static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// The namespace declared on the [Content_Types].xml
        /// </summary>
        [NotNull]
        protected static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        protected static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// The source file for this <see cref="OpenXmlVisitor"/>.
        /// </summary>
        [NotNull]
        public DocxFilePath File { get; }

        /// <summary>
        /// Active version of 'word/document.xml'.
        /// </summary>
        [NotNull]
        public virtual XElement Document { get; }

        /// <summary>
        /// Active version of 'word/_rels/document.xml.rels'.
        /// </summary>
        [NotNull]
        public virtual XElement DocumentRelations { get; }

        /// <summary>
        /// Active version of '[Content_Types].xml'.
        /// </summary>
        [NotNull]
        public virtual XElement ContentTypes { get; }

        /// <summary>
        /// Active version of 'word/footnotes.xml'.
        /// </summary>
        [NotNull]
        public virtual XElement Footnotes { get; }

        /// <summary>
        /// Active version of 'word/_rels/footnotes.xml.rels'.
        /// </summary>
        [NotNull]
        public virtual XElement FootnoteRelations { get; }

        /// <summary>
        /// Active version of word/charts/chart#.xml.
        /// </summary>
        [NotNull]
        public virtual IImmutableList<(string Name, XElement Chart)> Charts { get; }

        /// <summary>
        /// Returns the last document relation identifier in use by the container.
        /// </summary>
        public virtual int DocumentRelationId { get; }

        /// <summary>
        /// Returns the last footnote identifier currently in use by the container.
        /// </summary>
        public virtual int FootnoteId { get; }

        /// <summary>
        /// Returns the last footnote hyperlink identifier currently in use by the container.
        /// </summary>
        public virtual int FootnoteRelationId { get; }

        /// <summary>
        /// Initializes an <see cref="OpenXmlVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="result">The file to which changes can be saved.</param>
        public OpenXmlVisitor([NotNull] DocxFilePath result)
        {
            File = result;

            Document = 
                result.ReadAsXml() ?? throw new FileNotFoundException("document.xml");

            ContentTypes = 
                result.ReadAsXml("[Content_Types].xml") ?? throw new FileNotFoundException("[Content_Types].xml");

            XElement footnotes =
                Footnotes =
                    result.ReadAsXml("word/footnotes.xml") ?? new XElement(W + "footnotes");

            XElement documentRelations =
                DocumentRelations =
                    result.ReadAsXml("word/_rels/document.xml.rels") ?? new XElement(P + "Relationships");

            XElement footnoteRelations =
                FootnoteRelations =
                    result.ReadAsXml("word/_rels/footnotes.xml.rels") ?? new XElement(P + "Relationships");

            Charts =
                documentRelations.Elements()
                                  .Select(x => x.Attribute("Target")?.Value)
                                  .Where(x => x?.StartsWith("charts/") ?? false)
                                  .Select(x => (Name: x, Chart: result.ReadAsXml($"word/{x}")))
                                  .ToImmutableList();

            DocumentRelationId =
                documentRelations.Elements(P + "Relationship")
                                  .Attributes("Id")
                                  .Select(x => x.Value.ParseInt() ?? 0)
                                  .DefaultIfEmpty(0)
                                  .Max();

            FootnoteId =
                footnotes.Elements(W + "footnote")
                          .Attributes(W + "id")
                          .Select(x => x.Value.ParseInt() ?? 0)
                          .DefaultIfEmpty(0)
                          .Max();

            FootnoteRelationId =
                footnoteRelations.Elements(P + "Relationship")
                                  .Attributes("Id")
                                  .Select(x => x.Value.ParseInt() ?? 0)
                                  .DefaultIfEmpty(0)
                                  .Max();
        }

        /// <summary>
        /// Initializes an <see cref="Visitors.OpenXmlVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="document"></param>
        /// <param name="documentRelations"></param>
        /// <param name="contentTypes"></param>
        /// <param name="footnotes"></param>
        /// <param name="foonoteRelations"></param>
        /// <param name="charts"></param>
        /// <param name="currentFootnoteId"></param>
        /// <param name="currentFootnoteRelationId"></param>
        /// <param name="currentDocumentRelationId"></param>
        private OpenXmlVisitor([NotNull] DocxFilePath file, [NotNull] XElement document, XElement documentRelations, XElement contentTypes, XElement footnotes, XElement foonoteRelations, IEnumerable<(string Name, XElement Chart)> charts, int currentFootnoteId, int currentFootnoteRelationId, int currentDocumentRelationId)
        {
            File = file;
            Document = document.Clone();
            DocumentRelations = documentRelations.Clone();
            ContentTypes = contentTypes.Clone();
            Footnotes = footnotes.Clone();
            FootnoteRelations = foonoteRelations.Clone();
            Charts = charts.Select(x => (Name: x.Name, Chart: x.Chart.Clone())).ToImmutableArray();
            // TODO: Why isn't this needed? Actually, why does this proactively break the document?
            //_currentFootnoteId = currentFootnoteId;
            FootnoteRelationId = currentFootnoteRelationId;
            DocumentRelationId = currentDocumentRelationId;
        }

        /// <summary>
        /// Saves any modifications to <paramref name="resultPath"/>. This operation will overwrite any existing content for the modified parts.
        /// </summary>
        /// <param name="resultPath">The path to which modified parts are written.</param>
        public void Save([NotNull] DocxFilePath resultPath)
        {
            Document.WriteInto(resultPath, "word/document.xml");
            Footnotes.WriteInto(resultPath, "word/footnotes.xml");
            ContentTypes.WriteInto(resultPath, "[Content_Types].xml");
            DocumentRelations.WriteInto(resultPath, "word/_rels/document.xml.rels");
            FootnoteRelations.WriteInto(resultPath, "word/_rels/footnotes.xml.rels");
            foreach ((string Name, XElement Chart) chart in Charts)
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
        public OpenXmlVisitor Visit([NotNull][ItemNotNull] IEnumerable<DocxFilePath> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            return files.Aggregate(this, (current, next) => current.Visit(next));
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        [Pure]
        [NotNull]
        public OpenXmlVisitor Visit([NotNull] DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            OpenXmlVisitor subject = 
                new OpenXmlVisitor(file);

            OpenXmlDocumentVisitor contentVisitor =
                subject.VisitContent();

            OpenXmlFootnoteVisitor footnoteVisitor = 
                contentVisitor.VisitFootnotes(FootnoteId);

            OpenXmlFootnoteHyperlinkVisitor footnoteHyperlinkVisitor = 
                footnoteVisitor.VisitFootnoteHyperlinks(FootnoteRelationId);
            
            (XElement sourceContent3, XElement documentRelations1, int documentRelationId1) =
                file.MarshalContentHyperlinksFrom(footnoteVisitor.Document, DocumentRelationId);

            (XElement sourceContent4, XElement documentRelations2, XElement contentTypes1, IEnumerable<(string Name, XElement Chart)> charts1, int documentRelationId2) =
                file.MarshalChartsFrom(sourceContent3, ContentTypes, Charts, documentRelationId1);

            XElement resultContent =
                new XElement(
                    Document.Name,
                    Document.Attributes(),
                    new XElement(W + "body",
                        Document.Element(W + "body")?.Elements(),
                        sourceContent4.Element(W + "body")?.Elements()));

            XElement resultFootnotes =
                new XElement(
                    Footnotes.Name,
                    Footnotes.Attributes(),
                    Footnotes.Elements()
                             .Union(
                                 footnoteHyperlinkVisitor.Footnotes.Elements(),
                                 XNode.EqualityComparer));

            XElement resultFootnoteRelations =
                new XElement(
                    FootnoteRelations.Name,
                    FootnoteRelations.Attributes(),
                    FootnoteRelations.Elements()
                                     .Union(
                                         footnoteHyperlinkVisitor.FootnoteRelations.Elements(),
                                         XNode.EqualityComparer));

            XElement resultDocumentRelations =
                new XElement(
                    DocumentRelations.Name,
                    DocumentRelations.Attributes(),
                    DocumentRelations.Elements()
                                      .Union(
                                          documentRelations1?.Elements() ?? Enumerable.Empty<XElement>(),
                                          XNode.EqualityComparer)
                                      .Union(
                                          documentRelations2?.Elements() ?? Enumerable.Empty<XElement>(),
                                          XNode.EqualityComparer));

            return
                new OpenXmlVisitor(
                    file: File,
                    document: resultContent,
                    documentRelations: resultDocumentRelations,
                    contentTypes: contentTypes1,
                    footnotes: resultFootnotes,
                    foonoteRelations: resultFootnoteRelations,
                    charts: charts1,
                    currentFootnoteId: footnoteVisitor.FootnoteId,
                    currentFootnoteRelationId: footnoteHyperlinkVisitor.FootnoteRelationId,
                    currentDocumentRelationId: documentRelationId2);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public OpenXmlDocumentVisitor VisitContent()
        {
            return new OpenXmlDocumentVisitor(this);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="footnoteId"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public OpenXmlFootnoteVisitor VisitFootnotes(int footnoteId)
        {
            return new OpenXmlFootnoteVisitor(this, footnoteId);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="footnoteRelationId"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public OpenXmlFootnoteHyperlinkVisitor VisitFootnoteHyperlinks(int footnoteRelationId)
        {
            return new OpenXmlFootnoteHyperlinkVisitor(this, footnoteRelationId);
        }
    }
}