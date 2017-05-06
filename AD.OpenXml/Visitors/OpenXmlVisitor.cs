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
    ///   2) A private constructor that delegates to <see cref="OpenXmlVisitor(IOpenXmlVisitor)"/>.
    ///   3) Override <see cref="Create(IOpenXmlVisitor)"/>.
    ///   4) An optional override for each desired visitor method.
    /// </remarks>
    [PublicAPI]
    public class OpenXmlVisitor : IOpenXmlVisitor
    {
        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// word/charts/chart#.xml.
        /// </summary>
        public IEnumerable<ChartInformation> Charts { get; }

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        public XElement ContentTypes { get; }
        
        /// <summary>
        /// word/document.xml
        /// </summary>
        public XElement Document { get; }

        /// <summary>
        /// word/_rels/document.xml.rels
        /// </summary>
        public XElement DocumentRelations { get; }

        /// <summary>
        /// word/_rels/footnotes.xml.rels
        /// </summary>
        public XElement FootnoteRelations { get; }

        /// <summary>
        /// word/footnotes.xml
        /// </summary>
        public XElement Footnotes { get; }

        /// <summary>
        /// word/styles.xml
        /// </summary>
        public XElement Styles { get; }

        /// <summary>
        /// word/numbering.xml
        /// </summary>
        public XElement Numbering { get; }

        /// <summary>
        /// The current document relation number incremented by one.
        /// </summary>
        public int NextDocumentRelationId =>
            DocumentRelations.Elements().Count() + 1;

        /// <summary>
        /// The current footnote number incremented by one.
        /// </summary>
        public int NextFootnoteId =>
            Footnotes.Elements(W + "footnote")
                     .Count(x => int.Parse(x.Attribute(W + "id")?.Value ?? "0") > 0) + 1;

        /// <summary>
        /// The current footnote relation number incremented by one.
        /// </summary>
        public int NextFootnoteRelationId =>
            FootnoteRelations.Elements().Count() + 1;

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
                result.ReadAsXml("[Content_Types].xml") ?? throw new FileNotFoundException("[Content_Types].xml");

            Document = 
                result.ReadAsXml() ?? throw new FileNotFoundException("document.xml");
            
            DocumentRelations =
                result.ReadAsXml("word/_rels/document.xml.rels") ?? throw new FileNotFoundException("word/_rels/document.xml.rels");

            Footnotes =
                result.ReadAsXml("word/footnotes.xml") ?? new XElement(W + "footnotes");

            FootnoteRelations =
                result.ReadAsXml("word/_rels/footnotes.xml.rels") ?? new XElement(P + "Relationships");

            Styles =
                result.ReadAsXml("word/styles.xml") ?? throw new FileNotFoundException("word/styles.xml");

            Numbering =
                result.ReadAsXml("word/numbering.xml") ?? new XElement(W + "numbering");

            Charts =
                result.ReadAsXml("word/_rels/document.xml.rels")
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
        /// <param name="styles"></param>
        /// <param name="numbering"></param>
        /// <param name="charts">
        /// 
        /// </param>
        public OpenXmlVisitor([NotNull] XElement contentTypes, [NotNull] XElement document, [NotNull] XElement documentRelations, [NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, [NotNull] XElement styles, [NotNull] XElement numbering, [NotNull] IEnumerable<ChartInformation> charts)
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

        /// <summary>
        /// Writes the <see cref="IOpenXmlVisitor"/> to the <see cref="DocxFilePath"/>.
        /// </summary>
        /// <param name="result">
        /// The file to which the <see cref="IOpenXmlVisitor"/> is written.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public void Save(DocxFilePath result)
        {
            if (result is null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            Document.WriteInto(result, "word/document.xml");
            Footnotes.WriteInto(result, "word/footnotes.xml");
            ContentTypes.WriteInto(result, "[Content_Types].xml");
            DocumentRelations.WriteInto(result, "word/_rels/document.xml.rels");
            FootnoteRelations.WriteInto(result, "word/_rels/footnotes.xml.rels");
            Styles.WriteInto(result, "word/styles.xml");
            Numbering.WriteInto(result, "word/numbering.xml");
            foreach (ChartInformation item in Charts)
            {
                item.Chart.WriteInto(result, $"word/{item.Name}");
            }
        }

        /// <summary>
        /// Visit and fold the component documents into this <see cref="IOpenXmlVisitor"/>.
        /// </summary>
        /// <param name="files">
        /// The files to visit.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
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
        /// Folds <paramref name="subject"/> into this <see cref="IOpenXmlVisitor"/>.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="IOpenXmlVisitor"/> that is folded into this <see cref="IOpenXmlVisitor"/>.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        public virtual IOpenXmlVisitor Fold(IOpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(StaticFold(this, subject));
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
                              subject.Styles.Elements(),
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

            IEnumerable<ChartInformation> charts =
                source.Charts
                      .Union(
                          subject.Charts,
                          ChartInformation.Comparer);

            return new OpenXmlVisitor(contentTypes, document, documentRelations, footnotes, footnoteRelations, styles, numbering, charts);
        }

        /// <summary>
        /// Visit and join the component document into the <see cref="IOpenXmlVisitor"/>.
        /// </summary>
        /// <param name="file">
        /// The files to visit.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        public virtual IOpenXmlVisitor Visit(DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            IOpenXmlVisitor subject = new OpenXmlVisitor(file);
            IOpenXmlVisitor documentVisitor = VisitDocument(subject);
            IOpenXmlVisitor footnoteVisitor = VisitFootnotes(documentVisitor, NextFootnoteId);
            IOpenXmlVisitor documentRelationVisitor = VisitDocumentRelations(footnoteVisitor, NextDocumentRelationId);
            IOpenXmlVisitor footnoteRelationVisitor = VisitFootnoteRelations(documentRelationVisitor, NextFootnoteRelationId);
            IOpenXmlVisitor styleVisitor = VisitStyles(footnoteRelationVisitor);
            IOpenXmlVisitor numberingVisitor = VisitNumbering(styleVisitor);

            return numberingVisitor;
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
        protected virtual IOpenXmlVisitor VisitDocument([NotNull] IOpenXmlVisitor subject)
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
        protected virtual IOpenXmlVisitor VisitFootnotes([NotNull] IOpenXmlVisitor subject, int footnoteId)
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