﻿using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using AD.IO.Paths;
using AD.OpenXml.Structures;
using AD.OpenXml.Visits;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Represents a visitor or rewriter for OpenXML documents.
    /// </summary>
    /// <remarks>
    /// This class is modeled after the <see cref="T:System.Linq.Expressions.ExpressionVisitor" />.
    /// The goal is to encapsulate OpenXML manipulations within immutable objects. Every visit operation should be a pure function.
    /// Access to <see cref="T:System.Xml.Linq.XElement" /> objects should be done with care, ensuring that objects are cloned prior to any in-place mainpulations.
    /// The derived visitor class should provide:
    ///   1) A public constructor that delegates to <see cref="M:AD.OpenXml.OpenXmlPackageVisitor.#ctor(AD.IO.Paths.DocxFilePath)" />.
    ///   2) A private constructor that delegates to <see cref="M:AD.OpenXml.OpenXmlPackageVisitor.#ctor(AD.OpenXml.OpenXmlPackageVisitor)" />.
    ///   3) Override <see cref="M:AD.OpenXml.OpenXmlPackageVisitor.Create(AD.OpenXml.OpenXmlPackageVisitor)" />.
    ///   4) An optional override for each desired visitor method.
    /// </remarks>
    [PublicAPI]
    public sealed class OpenXmlPackageVisitor
    {
        [NotNull] private static readonly Package DefaultOpenXmlPackage =
            Package.Open(new MemoryStream(DocxFilePath.Create().ToArray()));

        [NotNull] private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public Document Document { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public Footnotes Footnotes { get; }

        /// <summary>
        /// word/styles.xml
        /// </summary>
        [NotNull]
        public XElement Styles { get; }

        /// <summary>
        /// word/numbering.xml
        /// </summary>
        [NotNull]
        public XElement Numbering { get; }

        /// <summary>
        /// word/theme/theme1.xml
        /// </summary>
        [NotNull]
        public XElement Theme1 { get; }

        /// <summary>
        /// The current revision number.
        /// </summary>
        public int RevisionId => Math.Max(Document.RevisionId, Footnotes.RevisionId);

        /// <summary>
        /// Initializes an <see cref="OpenXmlPackageVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="package">The package to which changes can be saved.</param>
        /// <exception cref="ArgumentNullException"/>
        public OpenXmlPackageVisitor([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            Document = new Document(package);
            Footnotes = new Footnotes(package);

            Uri stylesUri = new Uri("/word/styles.xml", UriKind.Relative);
            Uri numberingUri = new Uri("/word/numbering.xml", UriKind.Relative);
            Uri themeUri = new Uri("/word/theme/theme1.xml", UriKind.Relative);

            Styles =
                package.PartExists(stylesUri)
                    ? XElement.Load(package.GetPart(stylesUri).GetStream())
                    : new XElement(W + "styles");

            Numbering =
                package.PartExists(numberingUri)
                    ? XElement.Load(package.GetPart(numberingUri).GetStream())
                    : new XElement(W + "numbering");

            Theme1 =
                package.PartExists(themeUri)
                    ? XElement.Load(package.GetPart(themeUri).GetStream())
                    : new XElement(A + "theme");
        }

        /// <summary>
        /// Initializes a new <see cref="OpenXmlPackageVisitor"/> from the supplied components.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="footnotes"></param>
        /// <param name="styles"></param>
        /// <param name="numbering"></param>
        /// <param name="theme1"></param>
        /// <exception cref="ArgumentNullException"></exception>
        private OpenXmlPackageVisitor(
            [NotNull] Document document,
            [NotNull] Footnotes footnotes,
            [NotNull] XElement styles,
            [NotNull] XElement numbering,
            [NotNull] XElement theme1)
        {
            if (document is null)
                throw new ArgumentNullException(nameof(document));

            if (footnotes is null)
                throw new ArgumentNullException(nameof(footnotes));

            if (styles is null)
                throw new ArgumentNullException(nameof(styles));

            if (numbering is null)
                throw new ArgumentNullException(nameof(numbering));

            if (theme1 is null)
                throw new ArgumentNullException(nameof(theme1));

            Document = document;
            Footnotes = footnotes;
            Styles = styles;
            Numbering = numbering;
            Theme1 = theme1;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="document"></param>
        /// <param name="footnotes"></param>
        /// <param name="styles"></param>
        /// <param name="numbering"></param>
        /// <param name="theme1"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public OpenXmlPackageVisitor With(
            [CanBeNull] Document document = default,
            [CanBeNull] Footnotes footnotes = default,
            [CanBeNull] XElement styles = default,
            [CanBeNull] XElement numbering = default,
            [CanBeNull] XElement theme1 = default)
            => new OpenXmlPackageVisitor(
                document ?? Document,
                footnotes ?? Footnotes,
                styles ?? Styles,
                numbering ?? Numbering,
                theme1 ?? Theme1);

        /// <summary>
        /// Visit and join the component document into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="package">The package to visit.</param>
        [Pure]
        [NotNull]
        public OpenXmlPackageVisitor Visit([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            return
                new OpenXmlPackageVisitor(package)
                    .VisitDoc(RevisionId)
                    .VisitFootnotes(Footnotes.Count, RevisionId)
                    .VisitDocRels(Document.RelationshipsMax)
                    .VisitFootnotesRels(Footnotes.RelationshipsMax)
                    .VisitStyles()
                    .VisitNumbering();
        }

        /// <summary>
        /// Visit and fold the component documents into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="packages">The packages to visit.</param>
        [Pure]
        [NotNull]
        public static OpenXmlPackageVisitor VisitAndFold([NotNull] [ItemNotNull] IEnumerable<Package> packages)
        {
            if (packages is null)
                throw new ArgumentNullException(nameof(packages));

            return packages.Aggregate(new OpenXmlPackageVisitor(DefaultOpenXmlPackage), (current, next) => current.Fold(current.Visit(next)));
        }

        /// <summary>
        /// Folds <paramref name="subject"/> into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="subject">The visitor that is folded into the current visitor.</param>
        [Pure]
        [NotNull]
        private OpenXmlPackageVisitor Fold([NotNull] OpenXmlPackageVisitor subject)
        {
            if (subject is null)
                throw new ArgumentNullException(nameof(subject));

            Document document = Document.Concat(subject.Document);

            Footnotes footnotes = Footnotes.Concat(subject.Footnotes);

            XElement styles =
                new XElement(
                    Styles.Name,
                    Styles.Attributes(),
                    Styles.Elements()
                          .Union(
                              subject.Styles
                                     .Elements()
                                     .Where(x => x.Name != W + "docDefaults")
                                     .Where(x => (string) x.Attribute(W + "styleId") != "Normal"),
                              XNode.EqualityComparer));

            XElement numbering =
                new XElement(
                    Numbering.Name,
                    Numbering.Attributes(),
                    Numbering.Elements()
                             .Union(
                                 subject.Numbering.Elements(),
                                 XNode.EqualityComparer));

            // TODO: write a ThemeVisit
//            XElement theme1 =
//                new XElement(
//                    Theme1.Target,
//                    Theme1.Attributes(),
//                    Theme1.Elements()
//                          .Union(
//                              subject.Theme1.Elements(),
//                              XNode.EqualityComparer));

            return With(document, footnotes, styles, numbering, subject.Theme1);
        }

        /// <summary>
        /// Creates a <see cref="Package"/> from the <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="Package"/> representing the OpenXML package.
        /// </returns>
        [Pure]
        [NotNull]
        public Package ToPackage()
        {
            MemoryStream ms = new MemoryStream();
            ms.Write(DocxFilePath.Create().Span);

            Package package = Package.Open(ms, FileMode.Open);

            Document.Save(package);
            Footnotes.Save(package);

            BuildDocumentRelationships().Save(package, DocumentRelsInfo.PartName);
            BuildFootnoteRelationships().Save(package, FootnotesRelsInfo.PartName);

            Uri stylesUri = new Uri("/word/styles.xml", UriKind.Relative);
            using (Stream stream =
                package.PartExists(stylesUri)
                    ? package.GetPart(stylesUri).GetStream()
                    : package.CreatePart(stylesUri, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml").GetStream())
            {
                Styles.Save(stream);
            }

            Uri numberingUri = new Uri("/word/numbering.xml", UriKind.Relative);
            using (Stream stream =
                package.PartExists(numberingUri)
                    ? package.GetPart(numberingUri).GetStream()
                    : package.CreatePart(numberingUri, "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml").GetStream())
            {
                Numbering.Save(stream);
            }

            Uri themeUri = new Uri("/word/theme/theme1.xml", UriKind.Relative);
            using (Stream stream =
                package.PartExists(themeUri)
                    ? package.GetPart(themeUri).GetStream()
                    : package.CreatePart(themeUri, "application/vnd.openxmlformats-officedocument.theme+xml").GetStream())
            {
                Theme1.Save(stream);
            }

            return package;
        }

        /// <summary>
        /// Cosntructs a <see cref="Relationships"/> instance for /word/_rels/document.xml.rels.
        /// </summary>
        [Pure]
        [NotNull]
        private Relationships BuildDocumentRelationships()
            => new Relationships(
                new Relationships.Entry[]
                {
                    Footnotes.RelationshipEntry,
                    new Relationships.Entry("rId2", "numbering.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                    new Relationships.Entry("rId4", "styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
                    new Relationships.Entry("rId5", "theme/theme1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
                },
                Document.Charts.Select(x => x.RelationshipEntry),
                Document.Images.Select(x => x.RelationshipEntry),
                Document.Hyperlinks.Select(x => x.RelationshipEntry));

        /// <summary>
        /// Cosntructs a <see cref="Relationships"/> instance for /word/_rels/footnote.xml.rels.
        /// </summary>
        [Pure]
        [NotNull]
        private Relationships BuildFootnoteRelationships() => new Relationships(Footnotes.Hyperlinks.Select(x => x.RelationshipEntry));
    }
}