using System;
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
    /// Represents a visitor for OpenXML documents.
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
        [NotNull] private static readonly Package DefaultOpenXmlPackage = DocxFilePath.Create().ToPackage();
        [NotNull] private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;
        [NotNull] private static readonly XNamespace M = XNamespaces.OpenXmlMath;
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// The package that initialized the <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        [NotNull]
        public Package Package { get; }

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

            if (!package.FileOpenAccess.HasFlag(FileAccess.Read))
                throw new IOException("The package is write-only.");

            Package =
                package.FileOpenAccess.HasFlag(FileAccess.Write)
                    ? package.ToPackage()
                    : package;

            Document = new Document(package);
            Footnotes = new Footnotes(package);

            Uri stylesUri = new Uri("/word/styles.xml", UriKind.Relative);
            Styles =
                package.PartExists(stylesUri)
                    ? XElement.Load(package.GetPart(stylesUri).GetStream())
                    : new XElement(W + "styles",
                        new XAttribute(XNamespace.Xmlns + "w", W));

            Uri numberingUri = new Uri("/word/numbering.xml", UriKind.Relative);
            Numbering =
                package.PartExists(numberingUri)
                    ? XElement.Load(package.GetPart(numberingUri).GetStream())
                    : new XElement(W + "numbering",
                        new XAttribute(XNamespace.Xmlns + "w", W));

            Uri themeUri = new Uri("/word/theme/theme1.xml", UriKind.Relative);
            Theme1 =
                package.PartExists(themeUri)
                    ? XElement.Load(package.GetPart(themeUri).GetStream())
                    : new XElement(A + "theme",
                        new XAttribute(XNamespace.Xmlns + "a", A));
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
        {
            Package package = DefaultOpenXmlPackage.ToPackage(FileAccess.ReadWrite);

            (document ?? Document).CopyTo(package);
            (footnotes ?? Footnotes).CopyTo(package);

            PackagePart documentPart = package.GetPart(Document.PartUri);

            SaveHelper(
                package,
                documentPart,
                new Uri("/word/styles.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                styles ?? Styles);

            SaveHelper(
                package,
                documentPart,
                new Uri("/word/numbering.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
                numbering ?? Numbering);

            SaveHelper(
                package,
                documentPart,
                new Uri("/word/theme/theme1.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.theme+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                theme1 ?? Theme1);

            SaveHelper(
                package,
                documentPart,
                new Uri("/word/settings.xml", UriKind.Relative),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
                new XElement(W + "settings",
                    new XAttribute(XNamespace.Xmlns + "w", W),
                    new XAttribute(XNamespace.Xmlns + "m", M),
                    new XElement(W + "evenAndOddHeaders"),
                    new XElement(M + "mathPr",
                        new XElement(M + "mathFont",
                            new XAttribute(M + "val", "Cambria Math")),
                        new XElement(M + "intLim",
                            new XAttribute(M + "val", "subSup")),
                        new XElement(M + "naryLim",
                            new XAttribute(M + "val", "subSup"))),
                    new XElement(W + "rsids",
                        new XElement(W + "rsidRoot",
                            new XAttribute(W + "val", package.GetHashCode().ToString("X8"))))));

            return new OpenXmlPackageVisitor(package);
        }

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
                   .VisitStyles()
                   .VisitNumbering();
        }

        /// <summary>
        /// Visit and fold the component documents into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="packages">The packages to visit.</param>
        [Pure]
        [NotNull]
        public static OpenXmlPackageVisitor Visit([NotNull] [ItemCanBeNull] IEnumerable<Package> packages)
        {
            if (packages is null)
                throw new ArgumentNullException(nameof(packages));

            return packages.Where(x => x != null)
                           .Aggregate(
                                new OpenXmlPackageVisitor(DefaultOpenXmlPackage),
                                (current, next) => current.Fold(current.Visit(next)));
        }

        /// <summary>
        /// Folds <paramref name="subject"/> into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="subject">The visitor that is folded into the current visitor.</param>
        [Pure]
        [NotNull]
        private OpenXmlPackageVisitor Fold([NotNull] OpenXmlPackageVisitor subject)
        {
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
//            XElement theme =
//                new XElement(
//                    Theme.TargetUri,
//                    Theme.Attributes(),
//                    Theme.Elements()
//                          .Union(
//                              subject.Theme.Elements(),
//                              XNode.EqualityComparer));

            return With(document, footnotes, styles, numbering, subject.Theme1);
        }

        private static void SaveHelper(
            [NotNull] Package package,
            [NotNull] PackagePart owner,
            [NotNull] Uri uri,
            [NotNull] string contentType,
            [NotNull] string relationshipType,
            [NotNull] XElement element)
        {
            if (package.PartExists(uri))
                package.DeletePart(uri);

            owner.CreateRelationship(uri, TargetMode.Internal, relationshipType);

            PackagePart part = package.CreatePart(uri, contentType);

            element.WriteTo(part);
        }
    }
}