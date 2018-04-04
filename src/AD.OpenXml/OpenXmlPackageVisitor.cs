using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
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

        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

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

        [NotNull] private static readonly IEnumerable<string> UpdatableRelationTypes =
            new string[]
            {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
            };

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        [NotNull]
        public ContentTypes ContentTypes =>
            ContentTypes.Create(
                new ContentTypes.Override[]
                {
                    new ContentTypes.Override("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
                    new ContentTypes.Override("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
                    new ContentTypes.Override("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
                    new ContentTypes.Override("/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"),
                    new ContentTypes.Override("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"),
                    new ContentTypes.Override("/word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
                    new ContentTypes.Override("/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"),
                    new ContentTypes.Override("/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"),
                },
                Document.Charts.Select(x => x.ContentTypeEntry));

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public Document Document { get; }

        /// <summary>
        /// word/_rels/document.xml.rels
        /// </summary>
        [NotNull]
        public XElement DocumentRelations { get; }

        /// <summary>
        /// word/_rels/footnotes.xml.rels
        /// </summary>
        [NotNull]
        public XElement FootnoteRelations { get; }

        /// <summary>
        /// word/footnotes.xml
        /// </summary>
        [NotNull]
        public XElement Footnotes { get; }

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
        /// The current document relation count.
        /// </summary>
        public uint DocumentRelationCount => DocumentRelations.Elements().Max(x => uint.Parse(x.Attribute("Id").Value.Substring(3)));

        /// <summary>
        /// The current footnote count.
        /// </summary>
        public uint FootnoteCount => (uint) Footnotes.Elements().Skip(2).Count();

        /// <summary>
        /// The current footnote relation count.
        /// </summary>
        public uint FootnoteRelationCount => (uint) FootnoteRelations.Elements().Count();

        /// <summary>
        /// The current revision number.
        /// </summary>
        public uint RevisionId =>
            (uint) Math.Max(
                Document.Content.Descendants()
                        .Where(x => Revisions.Contains(x.Name))
                        .Select(x => (int) x.Attribute(W + "id"))
                        .DefaultIfEmpty(0)
                        .Max(),
                Footnotes.Descendants()
                         .Where(x => Revisions.Contains(x.Name))
                         .Select(x => (int) x.Attribute(W + "id"))
                         .DefaultIfEmpty(0)
                         .Max());

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
            Footnotes = archive.ReadXml("word/footnotes.xml", new XElement(W + "footnotes"));
            FootnoteRelations = archive.ReadXml("word/_rels/footnotes.xml.rels", new XElement(P + "Relationships"));
            Styles = archive.ReadXml("word/styles.xml");
            Numbering = archive.ReadXml("word/numbering.xml", new XElement(W + "numbering"));
            Theme1 = archive.ReadXml("word/theme/theme1.xml");

            XElement documentRelations = archive.ReadXml(DocumentRelsInfo.Path);

            DocumentRelations =
                new XElement(
                    P + "Relationships",
                    documentRelations.Elements().Where(x => !UpdatableRelationTypes.Contains(x.Attribute("Type").Value)),
                    Document.Charts.Select(x => (XElement) x.RelationshipEntry),
                    Document.Images.Select(x => (XElement) x.RelationshipEntry),
                    Document.Hyperlinks.Select(x => (XElement) x.RelationshipEntry));

            // ReSharper disable once InvertIf
            if (!Numbering.HasElements)
            {
                DocumentRelations =
                    new XElement(
                        DocumentRelations.Name,
                        DocumentRelations.Attributes(),
                        DocumentRelations.Elements().Where(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Target) != "numbering.xml"),
                        new XElement(
                            P + "Relationship",
                            new XAttribute("Id", $"rId{DocumentRelationCount + 1}"),
                            new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                            new XAttribute("Target", "numbering.xml")));
            }
        }

        /// <summary>
        /// Initializes a new <see cref="OpenXmlPackageVisitor"/> from the supplied components.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="documentRelations"></param>
        /// <param name="footnotes"></param>
        /// <param name="footnoteRelations"></param>
        /// <param name="styles"></param>
        /// <param name="numbering"></param>
        /// <param name="theme1"></param>
        /// <exception cref="ArgumentNullException"></exception>
        private OpenXmlPackageVisitor(
            [NotNull] Document document,
            [NotNull] XElement documentRelations,
            [NotNull] XElement footnotes,
            [NotNull] XElement footnoteRelations,
            [NotNull] XElement styles,
            [NotNull] XElement numbering,
            [NotNull] XElement theme1)
        {
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

            Document = document;
            Footnotes = footnotes;
            FootnoteRelations = footnoteRelations;
            Styles = styles;
            Numbering = numbering;
            Theme1 = theme1;
            DocumentRelations =
                new XElement(
                    P + "Relationships",
                    documentRelations.Elements().Where(x => !UpdatableRelationTypes.Contains(x.Attribute("Type").Value)),
                    Document.Charts.Select(x => (XElement) x.RelationshipEntry),
                    Document.Images.Select(x => (XElement) x.RelationshipEntry),
                    Document.Hyperlinks.Select(x => (XElement) x.RelationshipEntry));
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="document"></param>
        /// <param name="footnotes"></param>
        /// <param name="footnoteRelations"></param>
        /// <param name="styles"></param>
        /// <param name="numbering"></param>
        /// <param name="theme1"></param>
        /// <returns></returns>
        public OpenXmlPackageVisitor With(
            [CanBeNull] Document document = default,
            [CanBeNull] XElement footnotes = default,
            [CanBeNull] XElement footnoteRelations = default,
            [CanBeNull] XElement styles = default,
            [CanBeNull] XElement numbering = default,
            [CanBeNull] XElement theme1 = default)
        {
            return new OpenXmlPackageVisitor(
                document ?? Document,
                DocumentRelations,
                footnotes ?? Footnotes,
                footnoteRelations ?? FootnoteRelations,
                styles ?? Styles,
                numbering ?? Numbering,
                theme1 ?? Theme1);
        }

        /// <summary>
        /// Writes the <see cref="OpenXmlPackageVisitor"/> to the <see cref="DocxFilePath"/>.
        /// </summary>
        /// <returns>
        /// The stream to which the <see cref="DocxFilePath"/> is written.
        /// </returns>
        public ZipArchive Save()
        {
            ZipArchive archive = new ZipArchive(DocxFilePath.Create(), ZipArchiveMode.Update);

            Document.Save(archive);

            // TODO: remove when fully encapsulated
            using (Stream stream = archive.GetEntry(DocumentRelsInfo.Path).Open())
            {
                DocumentRelations.Save(stream);
            }

            using (Stream stream = archive.GetEntry("word/footnotes.xml").Open())
            {
                Footnotes.Save(stream);
            }

            using (Stream stream = archive.GetEntry(ContentTypesInfo.Path).Open())
            {
                using (StreamWriter writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    writer.Write(ContentTypes.ToString());
                }
            }

            using (Stream stream = archive.GetEntry("word/_rels/footnotes.xml.rels").Open())
            {
                FootnoteRelations.Save(stream);
            }

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
        /// Visit and join the component document into this <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="archive">
        /// The archive to visit.
        /// </param>
        [Pure]
        public OpenXmlPackageVisitor Visit(ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            OpenXmlPackageVisitor subject = new OpenXmlPackageVisitor(archive);
            OpenXmlPackageVisitor documentVisitor = new DocumentVisit(subject, RevisionId).Result;
            OpenXmlPackageVisitor footnoteVisitor = new FootnoteVisit(documentVisitor, FootnoteCount, RevisionId).Result;
            OpenXmlPackageVisitor documentRelationVisitor = new DocumentRelationVisit(footnoteVisitor, DocumentRelationCount).Result;
            OpenXmlPackageVisitor footnoteRelationVisitor = new FootnoteRelationVisit(documentRelationVisitor, FootnoteRelationCount).Result;
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
        public static OpenXmlPackageVisitor VisitAndFold(IEnumerable<ZipArchive> archives)
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
        public OpenXmlPackageVisitor Fold(OpenXmlPackageVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            Document document = Document.Concat(subject.Document);

            XElement footnotes =
                new XElement(
                    Footnotes.Name,
                    Footnotes.Attributes(),
                    Footnotes.Elements(),
                    subject.Footnotes.Elements());

            XElement footnoteRelations =
                new XElement(
                    FootnoteRelations.Name,
                    FootnoteRelations.Attributes(),
                    FootnoteRelations.Elements(),
                    subject.FootnoteRelations
                           .Elements()
                           .Where(x => UpdatableRelationTypes.Contains((string) x.Attribute("Type"))));

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
                    footnoteRelations: footnoteRelations,
                    styles: styles,
                    numbering: numbering,
                    theme1: subject.Theme1);
        }
    }
}