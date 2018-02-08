using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using AD.IO;
using AD.IO.Paths;
using AD.OpenXml.Elements;
using AD.OpenXml.Structures;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <inheritdoc />
    ///  <summary>
    ///  Represents a visitor or rewriter for OpenXML documents.
    ///  </summary>
    ///  <remarks>
    ///  This class is modeled after the <see cref="T:System.Linq.Expressions.ExpressionVisitor" />.
    ///  The goal is to encapsulate OpenXML manipulations within immutable objects. Every visit operation should be a pure function.
    ///  Access to <see cref="T:System.Xml.Linq.XElement" /> objects should be done with care, ensuring that objects are cloned prior to any in-place mainpulations.
    ///  The derived visitor class should provide:
    ///    1) A public constructor that delegates to <see cref="M:AD.OpenXml.Visitors.OpenXmlVisitor.#ctor(AD.IO.Paths.DocxFilePath)" />.
    ///    2) A private constructor that delegates to <see cref="M:AD.OpenXml.Visitors.OpenXmlVisitor.#ctor(AD.OpenXml.IOpenXmlVisitor)" />.
    ///    3) Override <see cref="M:AD.OpenXml.Visitors.OpenXmlVisitor.Create(AD.OpenXml.IOpenXmlVisitor)" />.
    ///    4) An optional override for each desired visitor method.
    ///  </remarks>
    [PublicAPI]
    public class OpenXmlVisitor : IOpenXmlVisitor
    {
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] private static readonly IEnumerable<XName> Revisions =
            new XName[]
            {
                W + "ins",
                W + "del",
                W + "rPrChange",
                W + "moveToRangeStart",
                W + "moveToRangeEnd",
                W + "moveTo"
            };

        /// <inheritdoc />
        public IEnumerable<ChartInformation> Charts { get; }

        /// <inheritdoc />
        public XElement ContentTypes { get; }

        /// <inheritdoc />
        public XElement Document { get; }

        /// <inheritdoc />
        public XElement DocumentRelations { get; }

        /// <inheritdoc />
        public XElement FootnoteRelations { get; }

        /// <inheritdoc />
        public XElement Footnotes { get; }

        /// <inheritdoc />
        public XElement Styles { get; }

        /// <inheritdoc />
        public XElement Numbering { get; }

        /// <inheritdoc />
        public XElement Theme1 { get; }

        /// <inheritdoc />
        public int NextDocumentRelationId => DocumentRelations.Elements().Count() + 1;

        /// <inheritdoc />
        public int NextFootnoteId => Footnotes.Elements(W + "footnote").Count(x => (int) x.Attribute(W + "id") > 0) + 1;

        /// <inheritdoc />
        public int NextFootnoteRelationId => FootnoteRelations.Elements().Count() + 1;

        /// <inheritdoc />
        public int NextRevisionId =>
            Math.Max(
                Document.Descendants()
                        .Where(x => Revisions.Contains(x.Name))
                        .Select(x => (int) x.Attribute(W + "id"))
                        .DefaultIfEmpty(0)
                        .Max(),
                Footnotes.Descendants()
                         .Where(x => Revisions.Contains(x.Name))
                         .Select(x => (int) x.Attribute(W + "id"))
                         .DefaultIfEmpty(0)
                         .Max()) + 1;

        /// <inheritdoc />
        /// <summary>
        /// Initializes an <see cref="T:AD.OpenXml.Visitors.OpenXmlVisitor" /> by reading document parts into memory from a default <see cref="MemoryStream"/>.
        /// </summary>
        public OpenXmlVisitor() : this(DocxFilePath.Create())
        {
        }

        /// <summary>
        /// Initializes an <see cref="OpenXmlVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="stream">
        /// The stream to which changes can be saved.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public OpenXmlVisitor([NotNull] MemoryStream stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            ContentTypes =
                stream.ReadAsXml(ContentTypesInfo.Path) ?? throw new FileNotFoundException(ContentTypesInfo.Path);

            Document =
                stream.ReadAsXml() ?? throw new FileNotFoundException("word/document.xml");

            DocumentRelations =
                stream.ReadAsXml(DocumentRelsInfo.Path) ?? throw new FileNotFoundException(DocumentRelsInfo.Path);

            Footnotes =
                stream.ReadAsXml("word/footnotes.xml") ?? new XElement(W + "footnotes");

            FootnoteRelations =
                stream.ReadAsXml("word/_rels/footnotes.xml.rels") ?? new XElement(P + "Relationships");

            Styles =
                stream.ReadAsXml("word/styles.xml") ?? throw new FileNotFoundException("word/styles.xml");

            Numbering =
                stream.ReadAsXml("word/numbering.xml");

            Theme1 =
                stream.ReadAsXml("word/theme/theme1.xml");

            if (Numbering is null)
            {
                Numbering = new XElement(W + "numbering");

                DocumentRelations =
                    new XElement(
                        DocumentRelations.Name,
                        DocumentRelations.Attributes(),
                        DocumentRelations.Elements()
                                         .Where(x => !x.Attribute("Tartget")?.Value.Equals("numbering.xml", StringComparison.OrdinalIgnoreCase) ?? true),
                        new XElement(
                            P + "Relationship",
                            new XAttribute("Id", $"rId{NextDocumentRelationId}"),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                            new XAttribute("Target", "numbering.xml")));

                ContentTypes =
                    new XElement(
                        ContentTypes.Name,
                        ContentTypes.Attributes(),
                        ContentTypes.Elements()
                                    .Where(x => !x.Attribute("PartName")?.Value.Equals("/word/numbering.xml", StringComparison.OrdinalIgnoreCase) ?? true),
                        new XElement(T + "Override",
                            new XAttribute("PartName", "/word/numbering.xml"),
                            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")));
            }

            Charts =
                stream.ReadAsXml(DocumentRelsInfo.Path)
                      .Elements()
                      .Select(x => x.Attribute("Target")?.Value)
                      .Where(x => x?.StartsWith("charts/") ?? false)
                      .Select(x => new ChartInformation(x, stream.ReadAsXml($"word/{x}")))
                      .ToImmutableList();
        }

        /// <summary>
        /// Initializes an <see cref="OpenXmlVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="result">
        /// The file to which changes can be saved.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public OpenXmlVisitor([NotNull] DocxFilePath result)
        {
            if (result is null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            ContentTypes =
                result.ReadAsXml(ContentTypesInfo.Path) ?? throw new FileNotFoundException(ContentTypesInfo.Path);

            Document =
                result.ReadAsXml() ?? throw new FileNotFoundException("word/document.xml");

            DocumentRelations =
                result.ReadAsXml(DocumentRelsInfo.Path) ?? throw new FileNotFoundException(DocumentRelsInfo.Path);

            Footnotes =
                result.ReadAsXml("word/footnotes.xml") ?? new XElement(W + "footnotes");

            FootnoteRelations =
                result.ReadAsXml("word/_rels/footnotes.xml.rels") ?? new XElement(P + "Relationships");

            Styles =
                result.ReadAsXml("word/styles.xml") ?? throw new FileNotFoundException("word/styles.xml");

            Numbering =
                result.ReadAsXml("word/numbering.xml");

            Theme1 =
                result.ReadAsXml("word/theme/theme1.xml");

            if (Numbering is null)
            {
                Numbering = new XElement(W + "numbering");

                DocumentRelations =
                    new XElement(
                        DocumentRelations.Name,
                        DocumentRelations.Attributes(),
                        DocumentRelations.Elements()
                                         .Where(x => !x.Attribute("Tartget")?.Value.Equals("numbering.xml", StringComparison.OrdinalIgnoreCase) ?? true),
                        new XElement(
                            P + "Relationship",
                            new XAttribute("Id", $"rId{NextDocumentRelationId}"),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                            new XAttribute("Target", "numbering.xml")));

                ContentTypes =
                    new XElement(
                        ContentTypes.Name,
                        ContentTypes.Attributes(),
                        ContentTypes.Elements()
                                    .Where(x => !x.Attribute("PartName")?.Value.Equals("/word/numbering.xml", StringComparison.OrdinalIgnoreCase) ?? true),
                        new XElement(T + "Override",
                            new XAttribute("PartName", "/word/numbering.xml"),
                            new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")));
            }

            Charts =
                result.ReadAsXml(DocumentRelsInfo.Path)
                      .Elements()
                      .Select(x => x.Attribute("Target")?.Value)
                      .Where(x => x?.StartsWith("charts/") ?? false)
                      .Select(x => new ChartInformation(x, result.ReadAsXml($"word/{x}")))
                      .ToImmutableList();
        }

        /// <summary>
        /// Initializes a new <see cref="OpenXmlVisitor"/> from an existing <see cref="IOpenXmlVisitor"/>.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="IOpenXmlVisitor"/> to visit.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public OpenXmlVisitor([NotNull] IOpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            Document = subject.Document.Clone();
            DocumentRelations = subject.DocumentRelations.Clone();
            ContentTypes = subject.ContentTypes.Clone();
            Footnotes = subject.Footnotes.Clone();
            FootnoteRelations = subject.FootnoteRelations.Clone();
            Styles = subject.Styles.Clone();
            Numbering = subject.Numbering.Clone();
            Theme1 = subject.Theme1.Clone();
            Charts = subject.Charts.Select(x => new ChartInformation(x.Name, x.Chart.Clone())).ToImmutableArray();
        }

        /// <summary>
        /// Initializes a new <see cref="OpenXmlVisitor"/> from the supplied components.
        /// </summary>
        /// <param name="contentTypes">
        ///
        /// </param>
        /// <param name="document">
        ///
        /// </param>
        /// <param name="documentRelations">
        ///
        /// </param>
        /// <param name="footnotes">
        ///
        /// </param>
        /// <param name="footnoteRelations">
        ///
        /// </param>
        /// <param name="styles">
        ///
        /// </param>
        /// <param name="numbering">
        ///
        /// </param>
        /// <param name="theme1">
        ///
        /// </param>
        /// <param name="charts">
        ///
        /// </param>
        public OpenXmlVisitor([NotNull] XElement contentTypes, [NotNull] XElement document, [NotNull] XElement documentRelations, [NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, [NotNull] XElement styles, [NotNull] XElement numbering, [NotNull] XElement theme1, [NotNull] IEnumerable<ChartInformation> charts)
        {
            if (contentTypes is null)
            {
                throw new ArgumentNullException(nameof(contentTypes));
            }
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }
            if (documentRelations is null)
            {
                throw new ArgumentNullException(nameof(documentRelations));
            }
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }
            if (footnoteRelations is null)
            {
                throw new ArgumentNullException(nameof(footnoteRelations));
            }
            if (styles is null)
            {
                throw new ArgumentNullException(nameof(styles));
            }
            if (numbering is null)
            {
                throw new ArgumentNullException(nameof(numbering));
            }
            if (theme1 is null)
            {
                throw new ArgumentNullException(nameof(theme1));
            }
            if (charts is null)
            {
                throw new ArgumentNullException(nameof(charts));
            }

            ContentTypes = contentTypes.Clone();
            Document = document.Clone();
            DocumentRelations = documentRelations.Clone();
            Footnotes = footnotes.Clone();
            FootnoteRelations = footnoteRelations.Clone();
            Styles = styles.Clone();
            Numbering = numbering.Clone();
            Theme1 = theme1.Clone();
            Charts = charts.Select(x => new ChartInformation(x.Name, x.Chart.Clone())).ToImmutableArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="subject"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"/>
        protected virtual IOpenXmlVisitor Create([NotNull] IOpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new OpenXmlVisitor(subject);
        }

        /// <inheritdoc />
        public void Save(DocxFilePath result)
        {
            if (result is null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            Document.WriteInto(result, "word/document.xml");
            Footnotes.WriteInto(result, "word/footnotes.xml");
            ContentTypes.WriteInto(result, ContentTypesInfo.Path);
            DocumentRelations.WriteInto(result, DocumentRelsInfo.Path);
            FootnoteRelations.WriteInto(result, "word/_rels/footnotes.xml.rels");
            Styles.WriteInto(result, "word/styles.xml");
            Numbering.WriteInto(result, "word/numbering.xml");
            Theme1.WriteInto(result, "word/theme/theme1.xml");

            foreach (ChartInformation item in Charts)
            {
                item.Chart.WriteInto(result, $"word/{item.Name}");
            }
        }

        /// <inheritdoc />
        public async Task<MemoryStream> Save()
        {
            MemoryStream stream = DocxFilePath.Create();

            stream = await Document.WriteInto(stream, "word/document.xml");
            stream = await Footnotes.WriteInto(stream, "word/footnotes.xml");
            stream = await ContentTypes.WriteInto(stream, ContentTypesInfo.Path);
            stream = await DocumentRelations.WriteInto(stream, DocumentRelsInfo.Path);
            stream = await FootnoteRelations.WriteInto(stream, "word/_rels/footnotes.xml.rels");
            stream = await Styles.WriteInto(stream, "word/styles.xml");
            stream = await Numbering.WriteInto(stream, "word/numbering.xml");
            stream = await Theme1.WriteInto(stream, "word/theme/theme1.xml");

            foreach (ChartInformation item in Charts)
            {
                stream = await item.Chart.WriteInto(stream, $"word/{item.Name}");
            }

            return stream;
        }

        /// <inheritdoc />
        [Pure]
        public virtual IOpenXmlVisitor Visit(MemoryStream stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            IOpenXmlVisitor subject = new OpenXmlVisitor(stream);
            IOpenXmlVisitor documentVisitor = VisitDocument(subject, NextRevisionId);
            IOpenXmlVisitor footnoteVisitor = VisitFootnotes(documentVisitor, NextFootnoteId, NextRevisionId);
            IOpenXmlVisitor documentRelationVisitor = VisitDocumentRelations(footnoteVisitor, NextDocumentRelationId);
            IOpenXmlVisitor footnoteRelationVisitor = VisitFootnoteRelations(documentRelationVisitor, NextFootnoteRelationId);
            IOpenXmlVisitor styleVisitor = VisitStyles(footnoteRelationVisitor);
            IOpenXmlVisitor numberingVisitor = VisitNumbering(styleVisitor);

            return numberingVisitor;
        }

        /// <inheritdoc />
        [Pure]
        public virtual IOpenXmlVisitor Visit(DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            IOpenXmlVisitor subject = new OpenXmlVisitor(file);
            IOpenXmlVisitor documentVisitor = VisitDocument(subject, NextRevisionId);
            IOpenXmlVisitor footnoteVisitor = VisitFootnotes(documentVisitor, NextFootnoteId, NextRevisionId);
            IOpenXmlVisitor documentRelationVisitor = VisitDocumentRelations(footnoteVisitor, NextDocumentRelationId);
            IOpenXmlVisitor footnoteRelationVisitor = VisitFootnoteRelations(documentRelationVisitor, NextFootnoteRelationId);
            IOpenXmlVisitor styleVisitor = VisitStyles(footnoteRelationVisitor);
            IOpenXmlVisitor numberingVisitor = VisitNumbering(styleVisitor);

            return numberingVisitor;
        }

        /// <inheritdoc />
        [Pure]
        public virtual IOpenXmlVisitor Fold(IOpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(StaticFold(this, subject));
        }

        /// <inheritdoc />
        [Pure]
        public virtual IOpenXmlVisitor VisitAndFold(IEnumerable<MemoryStream> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            return files.Aggregate(this as IOpenXmlVisitor, (current, next) => current.Fold(current.Visit(next)));
        }

        /// <inheritdoc />
        [Pure]
        public virtual IOpenXmlVisitor VisitAndFold(IEnumerable<DocxFilePath> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            return files.Aggregate(this as IOpenXmlVisitor, (current, next) => current.Fold(current.Visit(next)));
        }

        /// <summary>
        /// Folds <paramref name="subject"/> into this <paramref name="source"/>.
        /// </summary>
        /// <param name="source">
        /// The <see cref="IOpenXmlVisitor"/> into which the <paramref name="subject"/> is folded.
        /// </param>
        /// <param name="subject">
        /// The <see cref="IOpenXmlVisitor"/> that is folded into the <paramref name="source"/>.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        private static OpenXmlVisitor StaticFold([NotNull] IOpenXmlVisitor source, [NotNull] IOpenXmlVisitor subject)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            XElement document =
                new XElement(
                    source.Document.Name,
                    source.Document.Attributes(),
                    new XElement(
                        source.Document.Elements().First().Name,
                        source.Document.Elements().First().Elements(),
                        subject.Document.Elements().First().Elements()));

            document = document.RemoveDuplicateSectionProperties();

            XElement footnotes =
                new XElement(
                    source.Footnotes.Name,
                    source.Footnotes.Attributes(),
                    source.Footnotes
                          .Elements()
                          .Union(
                              subject.Footnotes.Elements(),
                              XNode.EqualityComparer));

            XElement footnoteRelations =
                new XElement(
                    source.FootnoteRelations.Name,
                    source.FootnoteRelations.Attributes(),
                    source.FootnoteRelations
                          .Elements()
                          .Union(
                              subject.FootnoteRelations.Elements(),
                              XNode.EqualityComparer));

            XElement documentRelations =
                new XElement(
                    source.DocumentRelations.Name,
                    source.DocumentRelations.Attributes(),
                    source.DocumentRelations
                          .Elements()
                          .Union(
                              subject.DocumentRelations.Elements(),
                              XNode.EqualityComparer));

            XElement contentTypes =
                new XElement(
                    source.ContentTypes.Name,
                    source.ContentTypes.Attributes(),
                    source.ContentTypes
                          .Elements()
                          .Union(
                              subject.ContentTypes.Elements(),
                              XNode.EqualityComparer));

            XElement styles =
                new XElement(
                    source.Styles.Name,
                    source.Styles.Attributes(),
                    source.Styles
                          .Elements()
                          .Union(
                              subject.Styles
                                     .Elements()
                                     .Where(x => x.Name != W + "docDefaults")
                                     .Where(x => (string) x.Attribute(W + "styleId") != "Normal"),
                              XNode.EqualityComparer));

            XElement numbering =
                new XElement(
                    source.Numbering.Name,
                    source.Numbering.Attributes(),
                    source.Numbering
                          .Elements()
                          .Union(
                              subject.Numbering.Elements(),
                              XNode.EqualityComparer));

            // TODO: write a ThemeVisit
//            XElement theme1 =
//                new XElement(
//                    source.Theme1.Name,
//                    source.Theme1.Attributes(),
//                    source.Theme1
//                          .Elements()
//                          .Union(
//                              subject.Theme1.Elements(),
//                              XNode.EqualityComparer));

            IEnumerable<ChartInformation> charts =
                source.Charts
                      .Union(
                          subject.Charts,
                          ChartInformation.Comparer);

            return
                new OpenXmlVisitor(
                    contentTypes,
                    document,
                    documentRelations,
                    footnotes,
                    footnoteRelations,
                    styles,
                    numbering,
                    subject.Theme1,
                    charts);
        }

        /// <summary>
        /// Visit the <see cref="Document"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <param name="revisionId">
        /// The current revision number incremented by one.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        protected virtual IOpenXmlVisitor VisitDocument([NotNull] IOpenXmlVisitor subject, int revisionId)
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
        /// <param name="revisionId">
        /// The current revision number incremented by one.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        protected virtual IOpenXmlVisitor VisitFootnotes([NotNull] IOpenXmlVisitor subject, int footnoteId, int revisionId)
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
        protected virtual IOpenXmlVisitor VisitDocumentRelations([NotNull] IOpenXmlVisitor subject, int documentRelationId)
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
        protected virtual IOpenXmlVisitor VisitFootnoteRelations([NotNull] IOpenXmlVisitor subject, int footnoteRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(subject);
        }

        /// <summary>
        /// Visit the <see cref="Styles"/> of the subject.
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
        protected virtual IOpenXmlVisitor VisitStyles([NotNull] IOpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(subject);
        }

        /// <summary>
        /// Visit the <see cref="Numbering"/> of the subject.
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
        protected virtual IOpenXmlVisitor VisitNumbering([NotNull] IOpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(subject);
        }
    }
}