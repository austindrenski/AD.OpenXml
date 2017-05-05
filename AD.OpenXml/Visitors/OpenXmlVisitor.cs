using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// Represents a visitor or rewriter for OpenXML documents.
    /// </summary>
    /// <remarks>
    /// This class is modeled after the <see cref="System.Linq.Expressions.ExpressionVisitor"/>. 
    /// 
    /// The goal is to encapsulate OpenXML manipulations within immutable objects. Every visit operation should be a pure function.
    /// Access to <see cref="XElement"/> objects should be done with care, ensuring that objects are cloned prior to any in-place mainpulations.
    /// 
    /// The derived visitor class should provide:
    ///   1) A public constructor that delegates to <see cref="OpenXmlVisitor(DocxFilePath)"/>.
    ///   2) A private constructor that delegates to <see cref="OpenXmlVisitor(OpenXmlVisitor)"/>.
    ///   3) Override <see cref="Create(OpenXmlVisitor)"/>.
    ///   4) Override <see cref="Create(DocxFilePath, XElement, XElement, XElement, XElement, XElement, IEnumerable{ChartInformation})"/>.
    ///   5) An optional override for each desired visitor method.
    /// </remarks>
    [PublicAPI]
    public class OpenXmlVisitor
    {
        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// The source file for this <see cref="OpenXmlVisitor"/>.
        /// </summary>
        [NotNull]
        public DocxFilePath File { get; }

        /// <summary>
        /// Active version of 'word/document.xml'.
        /// </summary>
        [NotNull]
        public XElement Document { get; }

        /// <summary>
        /// Active version of 'word/_rels/document.xml.rels'.
        /// </summary>
        [NotNull]
        public XElement DocumentRelations { get; }

        /// <summary>
        /// Active version of '[Content_Types].xml'.
        /// </summary>
        [NotNull]
        public XElement ContentTypes { get; }

        /// <summary>
        /// Active version of 'word/footnotes.xml'.
        /// </summary>
        [NotNull]
        public XElement Footnotes { get; }

        /// <summary>
        /// Active version of 'word/_rels/footnotes.xml.rels'.
        /// </summary>
        [NotNull]
        public XElement FootnoteRelations { get; }

        /// <summary>
        /// Active version of word/charts/chart#.xml.
        /// </summary>
        [NotNull]
        public IEnumerable<ChartInformation> Charts { get; }

        /// <summary>
        /// Returns the last document relationship identifier in use by the container.
        /// </summary>
        public int NextDocumentRelationId =>
            DocumentRelations.Elements().Count() + 1;

        /// <summary>
        /// Returns the last footnote identifier currently in use by the container.
        /// </summary>
        public int NextFootnoteId =>
            Footnotes.Elements(W + "footnote")
                     .Count(x => int.Parse(x.Attribute(W + "id")?.Value ?? "0") > 0) + 1;

        /// <summary>
        /// Returns the last footnote relationship identifier currently in use by the container.
        /// </summary>
        public int NextFootnoteRelationId =>
            FootnoteRelations.Elements().Count() + 1;

        /// <summary>
        /// Initializes an <see cref="OpenXmlVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="result">
        /// The file to which changes can be saved.
        /// </param>
        public OpenXmlVisitor([NotNull] DocxFilePath result)
        {
            File = result;

            ContentTypes =
                result.ReadAsXml("[Content_Types].xml") ?? throw new FileNotFoundException("[Content_Types].xml");

            Document = 
                result.ReadAsXml() ?? throw new FileNotFoundException("document.xml");
            
            DocumentRelations =
                result.ReadAsXml("word/_rels/document.xml.rels") ?? throw new FileNotFoundException("word/_rels/document.xml.rels");

            Footnotes =
                result.ReadAsXml("word/footnotes.xml") ?? new XElement(W + "footnotes");

            FootnoteRelations =
                result.ReadAsXml("word/_rels/footnotes.xml.rels") ?? new XElement(P + "Relationships");

            Charts =
                result.ReadAsXml("word/_rels/document.xml.rels")
                      .Elements()
                      .Select(x => x.Attribute("Target")?.Value)
                      .Where(x => x?.StartsWith("charts/") ?? false)
                      .Select(x => new ChartInformation(x, result.ReadAsXml($"word/{x}")))
                      .ToImmutableList();
        }

        /// <summary>
        /// Initializes a new <see cref="OpenXmlVisitor"/> from an existing <see cref="OpenXmlVisitor"/>.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        public OpenXmlVisitor([NotNull] OpenXmlVisitor subject)
        {
            File = subject.File;
            Document = subject.Document.Clone();
            DocumentRelations = subject.DocumentRelations.Clone();
            ContentTypes = subject.ContentTypes.Clone();
            Footnotes = subject.Footnotes.Clone();
            FootnoteRelations = subject.FootnoteRelations.Clone();
            Charts = subject.Charts.Select(x => new ChartInformation(x.Name, x.Chart.Clone())).ToImmutableArray();
        }

        /// <summary>
        /// Initializes a new <see cref="OpenXmlVisitor"/> from the supplied components. 
        /// This constructor should only be called within the base class.
        /// </summary>
        /// <param name="file">
        /// 
        /// </param>
        /// <param name="document">
        /// 
        /// </param>
        /// <param name="documentRelations">
        /// 
        /// </param>
        /// <param name="contentTypes">
        /// 
        /// </param>
        /// <param name="footnotes">
        /// 
        /// </param>
        /// <param name="footnoteRelations">
        /// 
        /// </param>
        /// <param name="charts">
        /// 
        /// </param>
        public OpenXmlVisitor([NotNull] DocxFilePath file, [NotNull] XElement document, [NotNull] XElement documentRelations, [NotNull] XElement contentTypes, [NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, [NotNull] IEnumerable<ChartInformation> charts)
        {
            File = file;
            Document = document.Clone();
            DocumentRelations = documentRelations.Clone();
            ContentTypes = contentTypes.Clone();
            Footnotes = footnotes.Clone();
            FootnoteRelations = footnoteRelations.Clone();
            Charts = charts.Select(x => new ChartInformation(x.Name, x.Chart.Clone())).ToImmutableArray();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        /// <returns></returns>
        protected virtual OpenXmlVisitor Create([NotNull] OpenXmlVisitor subject)
        {
            return new OpenXmlVisitor(subject);
        }

        /// <summary>
        /// Initializes a new <see cref="OpenXmlVisitor"/> from the supplied components. 
        /// This constructor should only be called within the base class.
        /// </summary>
        /// <param name="file">
        /// 
        /// </param>
        /// <param name="document">
        /// 
        /// </param>
        /// <param name="documentRelations">
        /// 
        /// </param>
        /// <param name="contentTypes">
        /// 
        /// </param>
        /// <param name="footnotes">
        /// 
        /// </param>
        /// <param name="footnoteRelations">
        /// 
        /// </param>
        /// <param name="charts">
        /// 
        /// </param>
        protected virtual OpenXmlVisitor Create([NotNull] DocxFilePath file, [NotNull] XElement document, [NotNull] XElement documentRelations, [NotNull] XElement contentTypes, [NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, [NotNull] IEnumerable<ChartInformation> charts)
        {
            return
                new OpenXmlVisitor(
                    File,
                    document,
                    documentRelations,
                    contentTypes,
                    footnotes,
                    footnoteRelations,
                    charts);
        }

        /// <summary>
        /// Saves the current visitor to the <see cref="DocxFilePath"/>.
        /// </summary>
        /// <param name="resultPath">
        /// The path to which modified parts are written.
        /// </param>
        public void Save([NotNull] DocxFilePath resultPath)
        {
            Document.WriteInto(resultPath, "word/document.xml");
            Footnotes.WriteInto(resultPath, "word/footnotes.xml");
            ContentTypes.WriteInto(resultPath, "[Content_Types].xml");
            DocumentRelations.WriteInto(resultPath, "word/_rels/document.xml.rels");
            FootnoteRelations.WriteInto(resultPath, "word/_rels/footnotes.xml.rels");
            foreach (ChartInformation item in Charts)
            {
                item.Chart.WriteInto(resultPath, $"word/{item.Name}");
            }
        }

        /// <summary>
        /// Visit and join the component documents into the <see cref="OpenXmlVisitor"/>.
        /// </summary>
        /// <param name="files">
        /// The files to visit.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        public virtual OpenXmlVisitor Visit([NotNull][ItemNotNull] IEnumerable<DocxFilePath> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            return files.Aggregate(this, (current, next) => current.Visit(next));
        }

        /// <summary>
        /// Visit and join the component document into the <see cref="OpenXmlVisitor"/>.
        /// </summary>
        /// <param name="file">
        /// The files to visit.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        public virtual OpenXmlVisitor Visit([NotNull] DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            OpenXmlVisitor subject = new OpenXmlVisitor(file);
            OpenXmlVisitor documentVisitor = VisitDocument(subject);
            OpenXmlVisitor footnoteVisitor = VisitFootnotes(documentVisitor, NextFootnoteId);
            OpenXmlVisitor footnoteRelationVisitor = VisitFootnoteRelations(footnoteVisitor, NextFootnoteRelationId);
            OpenXmlVisitor documentRelationVisitor = VisitDocumentRelations(footnoteRelationVisitor, NextDocumentRelationId);

            XElement document =
                new XElement(
                    Document.Name,
                    Document.Attributes(),
                    new XElement(
                        Document.Elements().First().Name,
                        Document.Elements().First().Elements(),
                        documentRelationVisitor.Document.Elements().First().Elements()));

            document = document.RemoveDuplicateSectionProperties();

            XElement footnotes =
                new XElement(
                    Footnotes.Name,
                    Footnotes.Attributes(),
                    Footnotes.Elements()
                             .Union(
                                 documentRelationVisitor.Footnotes.Elements(),
                                 XNode.EqualityComparer));

            XElement footnoteRelations =
                new XElement(
                    FootnoteRelations.Name,
                    FootnoteRelations.Attributes(),
                    FootnoteRelations.Elements()
                                     .Union(
                                         documentRelationVisitor.FootnoteRelations.Elements(),
                                         XNode.EqualityComparer));

            XElement documentRelations =
                new XElement(
                    DocumentRelations.Name,
                    DocumentRelations.Attributes(),
                    DocumentRelations.Elements()
                                     .Union(
                                         documentRelationVisitor.DocumentRelations.Elements(),
                                         XNode.EqualityComparer));
            
            XElement contentTypes =
                new XElement(
                    ContentTypes.Name,
                    ContentTypes.Attributes(),
                    ContentTypes.Elements()
                                .Union(
                                    documentRelationVisitor.ContentTypes.Elements(),
                                    XNode.EqualityComparer));

            IEnumerable<ChartInformation> charts =
                Charts.Union(
                    documentRelationVisitor.Charts,
                    ChartInformation.Comparer);

            return
                Create(
                    File,
                    document,
                    documentRelations,
                    contentTypes,
                    footnotes,
                    footnoteRelations,
                    charts);
        }

        /// <summary>
        /// Visit the <see cref="Document"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        protected virtual OpenXmlVisitor VisitDocument([NotNull] OpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(subject);
        }

        /// <summary>
        /// Visit the <see cref="Footnotes"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <param name="footnoteId">
        /// The current footnote identifier.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        protected virtual OpenXmlVisitor VisitFootnotes([NotNull] OpenXmlVisitor subject, int footnoteId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(subject);
        }

        /// <summary>
        /// Visit the <see cref="Document"/> and <see cref="DocumentRelations"/> of the subject to modify hyperlinks in the main document.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <param name="documentRelationId">
        /// The current document relationship identifier.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        protected virtual OpenXmlVisitor VisitDocumentRelations([NotNull] OpenXmlVisitor subject, int documentRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(subject);
        }

        /// <summary>
        /// Visit the <see cref="Footnotes"/> and <see cref="FootnoteRelations"/> of the subject to modify hyperlinks in the main document.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <param name="footnoteRelationId">
        /// The current footnote relationship identifier.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        protected virtual OpenXmlVisitor VisitFootnoteRelations([NotNull] OpenXmlVisitor subject, int footnoteRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(subject);
        }
    }
}