using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
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
        [NotNull] private static readonly ZipArchive DefaultOpenXml = new ZipArchive(DocxFilePath.Create());

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
        /// <param name="archive">
        /// The archive to which changes can be saved.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public OpenXmlPackageVisitor([NotNull] ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            Document = new Document(archive);
            Footnotes = new Footnotes(archive);
            Styles = archive.ReadXml("word/styles.xml");
            Numbering = archive.ReadXml("word/numbering.xml", new XElement(W + "numbering"));
            Theme1 = archive.ReadXml("word/theme/theme1.xml");
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
            {
                throw new ArgumentNullException(nameof(document));
            }

            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
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
        {
            return
                new OpenXmlPackageVisitor(
                    document ?? Document,
                    footnotes ?? Footnotes,
                    styles ?? Styles,
                    numbering ?? Numbering,
                    theme1 ?? Theme1);
        }

        /// <summary>
        /// Visit and join the component document into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="archive">
        /// The archive to visit.
        /// </param>
        [Pure]
        [NotNull]
        public OpenXmlPackageVisitor Visit([NotNull] ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            OpenXmlPackageVisitor subject = new OpenXmlPackageVisitor(archive);
            OpenXmlPackageVisitor documentVisitor = new DocumentVisit(subject, RevisionId).Result;
            OpenXmlPackageVisitor footnoteVisitor = new FootnoteVisit(documentVisitor, Footnotes.Count, RevisionId).Result;
            OpenXmlPackageVisitor documentRelationVisitor = new DocumentRelationVisit(footnoteVisitor, Document.RelationshipsMax).Result;
            OpenXmlPackageVisitor footnoteRelationVisitor = new FootnoteRelationVisit(documentRelationVisitor, Footnotes.RelationshipsMax).Result;
            OpenXmlPackageVisitor styleVisitor = new StyleVisit(footnoteRelationVisitor).Result;
            OpenXmlPackageVisitor numberingVisitor = new NumberingVisit(styleVisitor).Result;

            return numberingVisitor;
        }

        /// <summary>
        /// Visit and fold the component documents into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="archives">
        /// The archives to visit.
        /// </param>
        [Pure]
        [NotNull]
        public static OpenXmlPackageVisitor VisitAndFold([NotNull] [ItemNotNull] IEnumerable<ZipArchive> archives)
        {
            if (archives is null)
            {
                throw new ArgumentNullException(nameof(archives));
            }

            return archives.Aggregate(new OpenXmlPackageVisitor(DefaultOpenXml), (current, next) => current.Fold(current.Visit(next)));
        }

        /// <summary>
        /// Folds <paramref name="subject"/> into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlPackageVisitor"/> that is folded into this <see cref="OpenXmlPackageVisitor"/>.
        /// </param>
        [Pure]
        [NotNull]
        public OpenXmlPackageVisitor Fold([NotNull] OpenXmlPackageVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

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

            return
                With(
                    document: document,
                    footnotes: footnotes,
                    styles: styles,
                    numbering: numbering,
                    theme1: subject.Theme1);
        }

        /// <summary>
        /// Creates a <see cref="ZipArchive"/> from the <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <returns>
        /// A <see cref="ZipArchive"/> representing the OpenXML package.
        /// </returns>
        [Pure]
        [NotNull]
        public ZipArchive ToZipArchive()
        {
            ZipArchive archive = new ZipArchive(DocxFilePath.Create(), ZipArchiveMode.Update);

            Document.Save(archive);
            Footnotes.Save(archive);

            BuildContentTypes().Save(archive);
            BuildDocumentRelationships().Save(archive, DocumentRelsInfo.Path);
            BuildFootnoteRelationships().Save(archive, FootnotesRelsInfo.Path);

            using (Stream stream = archive.GetEntry("word/styles.xml").Open())
            {
                Styles.Save(stream);
            }

            using (Stream stream = archive.GetEntry("word/numbering.xml")?.Open() ??
                                   archive.CreateEntry("word/numbering.xml").Open())
            {
                Numbering.Save(stream);
            }

            using (Stream stream = archive.GetEntry("word/theme/theme1.xml")?.Open() ??
                                   archive.CreateEntry("word/theme/theme1.xml").Open())
            {
                Theme1.Save(stream);
            }

            return archive;
        }

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        [Pure]
        [NotNull]
        private ContentTypes BuildContentTypes()
        {
            return
                new ContentTypes(
                    new ContentTypes.Override[]
                    {
                        new ContentTypes.Override("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
                        new ContentTypes.Override("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
                        new ContentTypes.Override("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
                        new ContentTypes.Override("/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"),
                        new ContentTypes.Override("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"),
                        new ContentTypes.Override("/word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
                        Footnotes.ContentTypeEntry,
                        new ContentTypes.Override("/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"),
                    },
                    Document.Charts.Select(x => x.ContentTypeEntry));
        }

        /// <summary>
        /// Cosntructs a <see cref="Relationships"/> instance for /word/_rels/document.xml.rels.
        /// </summary>
        [Pure]
        [NotNull]
        private Relationships BuildDocumentRelationships()
        {
            return
                new Relationships(
                    new Relationships.Entry[]
                    {
                        Footnotes.RelationshipEntry,
                        new Relationships.Entry("rId2", "numbering.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                        new Relationships.Entry("rId3", "settings.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"),
                        new Relationships.Entry("rId4", "styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
                        new Relationships.Entry("rId5", "theme/theme1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"),
                    },
                    Document.Charts.Select(x => x.RelationshipEntry),
                    Document.Images.Select(x => x.RelationshipEntry),
                    Document.Hyperlinks.Select(x => x.RelationshipEntry));
        }

        /// <summary>
        /// Cosntructs a <see cref="Relationships"/> instance for /word/_rels/footnote.xml.rels.
        /// </summary>
        [Pure]
        [NotNull]
        private Relationships BuildFootnoteRelationships()
        {
            return new Relationships(Footnotes.Hyperlinks.Select(x => x.RelationshipEntry));
        }
    }
}