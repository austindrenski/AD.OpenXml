using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.IO.Paths;
using AD.OpenXml.Elements;
using AD.OpenXml.Structures;
using AD.OpenXml.Visitors;
using AD.OpenXml.Visits;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <inheritdoc />
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
    public sealed class OpenXmlPackageVisitor : IDisposable
    {
        [NotNull] private static readonly ZipArchive DefaultOpenXml = new ZipArchive(DocxFilePath.Create());

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

        [NotNull] private static readonly XElement NumberingOverrideEntry =
            new XElement(T + "Override",
                new XAttribute("PartName", "/word/numbering.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"));

        [NotNull]
        private XElement NumberingTargetEntry =>
            new XElement(
                P + "Relationship",
                new XAttribute("Id", $"rId{DocumentRelationId + 1}"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                new XAttribute("Target", "numbering.xml"));

        /// <summary>
        /// word/charts/chart#.xml.
        /// </summary>
        [NotNull]
        public IEnumerable<ChartInformation> Charts { get; }

        /// <summary>
        /// word/media/image#.[jpeg|png|svg].
        /// </summary>
        [NotNull]
        public IEnumerable<ImageInformation> Images { get; }

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        [NotNull]
        public XElement ContentTypes { get; }

        /// <summary>
        /// word/document.xml
        /// </summary>
        [NotNull]
        public XElement Document { get; }

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
        /// The current document relation number incremented by one.
        /// </summary>
        public int DocumentRelationId => DocumentRelations.Elements().Count();

        /// <summary>
        /// The current footnote number incremented by one.
        /// </summary>
        public int FootnoteId => Footnotes.Elements(W + "footnote").Count(x => (int) x.Attribute(W + "id") > 0);

        /// <summary>
        /// The current footnote relation number incremented by one.
        /// </summary>
        public int FootnoteRelationId => FootnoteRelations.Elements().Count();

        /// <summary>
        /// The current revision number incremented by one.
        /// </summary>
        public int RevisionId =>
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
                         .Max());

        /// <summary>
        /// Maps chart reference id to chart node.
        /// </summary>
        [NotNull]
        public IDictionary<string, XElement> ChartReferences =>
            DocumentRelations.Elements()
                             .Where(x => x.Attribute("Target").Value.StartsWith("charts/"))
                             .ToDictionary(
                                 x => (string) x.Attribute("Id"),
                                 x => Charts.Single(y => y.Name == (string) x.Attribute("Target")).Chart);

        /// <summary>
        /// Maps image reference id to image node.
        /// </summary>
        [NotNull]
        public IDictionary<string, (string mime, string description, string base64)> ImageReferences =>
            DocumentRelations.Elements()
                             .Where(x => x.Attribute("Target").Value.StartsWith("media/"))
                             .Select(
                                 x => new
                                 {
                                     id = (string) x.Attribute("Id"),
                                     target = (string) x.Attribute("Target"),
                                     description = string.Empty,
                                     base64 =
                                         Convert.ToBase64String(
                                             Images.Single(y => y.Name == (string) x.Attribute("Target")).Image.Span.ToArray()),
                                 })
                             .ToDictionary(
                                 x => x.id,
                                 x => (mime: x.target.Substring(x.target.LastIndexOf('.') + 1), x.description, x.base64));

        /// <summary>
        /// Initializes an <see cref="OpenXmlPackageVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="archive">
        /// The archive to which changes can be saved.
        /// </param>
        /// <param name="dispose">
        /// True if the archive should be disposed; otherwise false.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public OpenXmlPackageVisitor([NotNull] ZipArchive archive, bool dispose = true)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            ContentTypes = archive.ReadXml(ContentTypesInfo.Path);
            Document = archive.ReadXml();
            DocumentRelations = archive.ReadXml(DocumentRelsInfo.Path);
            Footnotes = archive.ReadXml("word/footnotes.xml", new XElement(W + "footnotes"));
            FootnoteRelations = archive.ReadXml("word/_rels/footnotes.xml.rels", new XElement(P + "Relationships"));
            Styles = archive.ReadXml("word/styles.xml");
            Numbering = archive.ReadXml("word/numbering.xml", new XElement(W + "numbering"));
            Theme1 = archive.ReadXml("word/theme/theme1.xml");

            if (!Numbering.HasElements)
            {
                DocumentRelations =
                    new XElement(
                        DocumentRelations.Name,
                        DocumentRelations.Attributes(),
                        DocumentRelations.Elements().Where(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Target) != "numbering.xml"),
                        NumberingTargetEntry);

                ContentTypes =
                    new XElement(
                        ContentTypes.Name,
                        ContentTypes.Attributes(),
                        ContentTypes.Elements().Where(x => (string) x.Attribute(ContentTypesInfo.Attributes.PartName) != "/word/numbering.xml"),
                        NumberingOverrideEntry);
            }

            Charts =
                archive.ReadXml(DocumentRelsInfo.Path)
                       .Elements()
                       .Select(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Target))
                       .Where(x => x?.StartsWith("charts/") ?? false)
                       .Select(x => new ChartInformation(x, archive.ReadXml($"word/{x}")))
                       .ToArray();

            Images =
                archive.ReadXml(DocumentRelsInfo.Path)
                       .Elements()
                       .Select(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Target))
                       .Where(x => x?.StartsWith("media/") ?? false)
                       .Select(x => new ImageInformation(x, archive.ReadByteArray($"word/{x}")))
                       .ToArray();
        }

        /// <summary>
        /// Initializes a new <see cref="OpenXmlPackageVisitor"/> from an existing <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlPackageVisitor"/> to visit.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public OpenXmlPackageVisitor([NotNull] OpenXmlPackageVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            Document = subject.Document;
            DocumentRelations = subject.DocumentRelations;
            ContentTypes = subject.ContentTypes;
            Footnotes = subject.Footnotes;
            FootnoteRelations = subject.FootnoteRelations;
            Styles = subject.Styles;
            Numbering = subject.Numbering;
            Theme1 = subject.Theme1;
            Charts = subject.Charts;
            Images = subject.Images;
        }

        ///  <summary>
        ///  Initializes a new <see cref="OpenXmlPackageVisitor"/> from the supplied components.
        ///  </summary>
        ///  <param name="contentTypes">
        ///
        ///  </param>
        ///  <param name="document">
        ///
        ///  </param>
        ///  <param name="documentRelations">
        ///
        ///  </param>
        ///  <param name="footnotes">
        ///
        ///  </param>
        ///  <param name="footnoteRelations">
        ///
        ///  </param>
        ///  <param name="styles">
        ///
        ///  </param>
        ///  <param name="numbering">
        ///
        ///  </param>
        ///  <param name="theme1">
        ///
        ///  </param>
        ///  <param name="charts">
        ///
        ///  </param>
        /// <param name="images">
        ///
        /// </param>
        internal OpenXmlPackageVisitor([NotNull] XElement contentTypes, [NotNull] XElement document, [NotNull] XElement documentRelations, [NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, [NotNull] XElement styles, [NotNull] XElement numbering, [NotNull] XElement theme1, [NotNull] IEnumerable<ChartInformation> charts, [NotNull] IEnumerable<ImageInformation> images)
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

            if (images is null)
            {
                throw new ArgumentNullException(nameof(images));
            }

            ContentTypes = contentTypes;
            Document = document;
            DocumentRelations = documentRelations;
            Footnotes = footnotes;
            FootnoteRelations = footnoteRelations;
            Styles = styles;
            Numbering = numbering;
            Theme1 = theme1;
            Charts = charts.ToArray();
            Images = images.ToArray();
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

            using (Stream stream = archive.GetEntry("word/document.xml").Open())
            {
                Document.Save(stream);
            }

            using (Stream stream = archive.GetEntry("word/footnotes.xml").Open())
            {
                Footnotes.Save(stream);
            }

            using (Stream stream = archive.GetEntry(ContentTypesInfo.Path).Open())
            {
                ContentTypes.Save(stream);
            }

            using (Stream stream = archive.GetEntry(DocumentRelsInfo.Path).Open())
            {
                DocumentRelations.Save(stream);
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

            foreach (ChartInformation item in Charts)
            {
                using (Stream stream = archive.CreateEntry($"word/{item.Name}").Open())
                {
                    item.Chart.Save(stream);
                }
            }

            foreach (ImageInformation item in Images)
            {
                using (Stream stream = archive.CreateEntry($"word/{item.Name}").Open())
                {
                    for (int i = 0; i < item.Image.Span.Length; i++)
                    {
                        stream.WriteByte(item.Image.Span[i]);
                    }
                }
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
            OpenXmlPackageVisitor footnoteVisitor = new FootnoteVisit(documentVisitor, FootnoteId, RevisionId).Result;
            OpenXmlPackageVisitor documentRelationVisitor = new DocumentRelationVisit(footnoteVisitor, DocumentRelationId).Result;
            OpenXmlPackageVisitor footnoteRelationVisitor = new FootnoteRelationVisit(documentRelationVisitor, FootnoteRelationId).Result;
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

            XElement document =
                new XElement(
                    Document.Name,
                    Document.Attributes(),
                    new XElement(
                        W + "body",
                        Document.Element(W + "body").Elements(),
                        subject.Document.Element(W + "body").Elements()));

            document.RemoveDuplicateSectionProperties();

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
                    subject.FootnoteRelations.Elements());

            XElement documentRelations =
                new XElement(
                    DocumentRelations.Name,
                    DocumentRelations.Attributes(),
                    DocumentRelations.Elements()
                                     .Union(
                                         subject.DocumentRelations.Elements(),
                                         XNode.EqualityComparer));

            XElement contentTypes =
                new XElement(
                    ContentTypes.Name,
                    ContentTypes.Attributes(),
                    ContentTypes.Elements()
                                .Union(
                                    subject.ContentTypes.Elements(),
                                    XNode.EqualityComparer));

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
//                    Theme1.Name,
//                    Theme1.Attributes(),
//                    Theme1.Elements()
//                          .Union(
//                              subject.Theme1.Elements(),
//                              XNode.EqualityComparer));

            IEnumerable<ChartInformation> charts = Charts.Concat(subject.Charts);

            IEnumerable<ImageInformation> images = Images.Concat(subject.Images);

            return
                new OpenXmlPackageVisitor(
                    contentTypes,
                    document,
                    documentRelations,
                    footnotes,
                    footnoteRelations,
                    styles,
                    numbering,
                    subject.Theme1,
                    charts,
                    images);
        }

        /// <inheritdoc />
        public void Dispose()
        {
            ContentTypes.RemoveAll();
            Document.RemoveAll();
            DocumentRelations.RemoveAll();
            FootnoteRelations.RemoveAll();
            Footnotes.RemoveAll();
            Numbering.RemoveAll();
            Styles.RemoveAll();
            Theme1.RemoveAll();
        }
    }
}