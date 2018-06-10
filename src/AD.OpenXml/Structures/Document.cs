using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public class Document
    {
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

        /// <summary>
        /// The XML file located at: /word/document.xml.
        /// </summary>
        [NotNull]
        public XElement Content { get; }

        /// <summary>
        /// The XML files located in: /word/charts/.
        /// </summary>
        [NotNull]
        public IEnumerable<ChartInfo> Charts { get; }

        /// <summary>
        /// The XML files located in: /word/media/.
        /// </summary>
        [NotNull]
        public IEnumerable<ImageInfo> Images { get; }

        /// <summary>
        /// The hyperlinks listed in: /word/_rels/document.xml.rels.
        /// </summary>
        [NotNull]
        public IEnumerable<HyperlinkInfo> Hyperlinks { get; }

        /// <summary>
        /// The number of relationships in the document.
        /// </summary>
        public int RelationshipsMax =>
            Charts.Select(x => x.NumericId)
                  .Concat(Images.Select(x => x.NumericId))
                  .Concat(Hyperlinks.Select(x => x.NumericId))
                  .DefaultIfEmpty(default)
                  .Max();

        /// <summary>
        /// The current revision number.
        /// </summary>
        public int RevisionId =>
            Content.Descendants()
                   .Where(x => Revisions.Contains(x.Name))
                   .Select(x => (int) x.Attribute(W + "id"))
                   .DefaultIfEmpty(0)
                   .Max();

        /// <summary>
        /// Maps chart reference id to chart node.
        /// </summary>
        [NotNull]
        public IDictionary<string, XElement> ChartReferences => Charts.ToDictionary(x => x.RelationId, x => x.Chart);

        /// <summary>
        /// Maps image reference id to image node.
        /// </summary>
        [NotNull]
        public IDictionary<string, (string mime, string description, string base64)> ImageReferences =>
            Images.ToDictionary(x => x.RelationId.ToString(), x => (x.Extension.ToString(), string.Empty, x.Base64String));

        /// <summary>
        /// Initializes an <see cref="OpenXmlPackageVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="archive">
        /// The archive to which changes can be saved.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public Document([NotNull] ZipArchive archive)
        {
            if (archive is null)
                throw new ArgumentNullException(nameof(archive));

            Content = archive.ReadXml();

            XElement documentRelations = archive.ReadXml(DocumentRelsInfo.Path, Relationships.Empty.ToXElement());

            // TODO: re-enumerate ids at this stage.
            // TODO: find charts and images by *starting* with the content refs.
//            Content.ChangeXAttributeValues("Id", "", "");

            Charts =
                documentRelations
                    .Elements()
                    .Select(
                        x =>
                            new
                            {
                                Id = (string) x.Attribute(DocumentRelsInfo.Attributes.Id),
                                Target = (string) x.Attribute(DocumentRelsInfo.Attributes.Target)
                            })
                    .Where(x => x.Target.StartsWith("charts/"))
                    .Select(x => new ChartInfo(x.Id, archive.ReadXml($"word/{x.Target}")))
                    .ToArray();

            Images =
                documentRelations
                    .Elements()
                    .Select(
                        x =>
                            new
                            {
                                Id = (string) x.Attribute(DocumentRelsInfo.Attributes.Id),
                                Target = (string) x.Attribute(DocumentRelsInfo.Attributes.Target)
                            })
                    .Where(x => x.Target.StartsWith("media/"))
                    .Select(x => ImageInfo.Create(x.Id, x.Target, archive.ReadByteArray($"word/{x.Target}")))
                    .ToArray();

            Hyperlinks =
                documentRelations
                    .Elements()
                    .Select(
                        x =>
                            new
                            {
                                Id = (string) x.Attribute(DocumentRelsInfo.Attributes.Id),
                                Target = (string) x.Attribute(DocumentRelsInfo.Attributes.Target),
                                TargetMode = (string) x.Attribute("TargetMode")
                            })
                    .Where(x => x.TargetMode != null)
                    .Select(x => new HyperlinkInfo(x.Id, x.Target, x.TargetMode))
                    .ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="charts"></param>
        /// <param name="images"> </param>
        /// <param name="hyperlinks"></param>
        public Document(
            [NotNull] XElement content,
            [NotNull] IEnumerable<ChartInfo> charts,
            [NotNull] IEnumerable<ImageInfo> images,
            [NotNull] IEnumerable<HyperlinkInfo> hyperlinks)
        {
            if (content is null)
                throw new ArgumentNullException(nameof(content));

            if (charts is null)
                throw new ArgumentNullException(nameof(charts));

            if (images is null)
                throw new ArgumentNullException(nameof(images));

            if (hyperlinks is null)
                throw new ArgumentNullException(nameof(hyperlinks));

            Content = content.Clone();
            Charts = charts.ToArray();
            Images = images.ToArray();
            Hyperlinks = hyperlinks.ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="charts"></param>
        /// <param name="images"></param>
        /// <param name="hyperlinks"></param>
        /// <returns></returns>
        public Document With(
            [CanBeNull] XElement content = default,
            [CanBeNull] IEnumerable<ChartInfo> charts = default,
            [CanBeNull] IEnumerable<ImageInfo> images = default,
            [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
            => new Document(
                content ?? Content,
                charts ?? Charts,
                images ?? Images,
                hyperlinks ?? Hyperlinks);

        /// <summary>
        ///
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public Document Concat([NotNull] Document other)
            => Concat(other.Content, other.Charts, other.Images, other.Hyperlinks);

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="charts"></param>
        /// <param name="images"></param>
        /// <param name="hyperlinks"></param>
        /// <returns></returns>
        public Document Concat(
            [CanBeNull] XElement content = default,
            [CanBeNull] IEnumerable<ChartInfo> charts = default,
            [CanBeNull] IEnumerable<ImageInfo> images = default,
            [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
        {
            XElement document =
                content is null
                    ? Content
                    : new XElement(
                        Content.Name,
                        Content.Attributes()
                               .Concat(content.Attributes())
                               .Select(x => (x.Name, x.Value))
                               .Distinct()
                               .Select(x => new XAttribute(x.Name, x.Value)),
                        new XElement(
                            W + "body",
                            Content.Element(W + "body")?.Elements(),
                            content.Element(W + "body")?.Elements()));

            document.RemoveDuplicateSectionProperties();

            return
                new Document(
                    document,
                    charts is null ? Charts : Charts.Concat(charts),
                    images is null ? Images : Images.Concat(images),
                    hyperlinks is null ? Hyperlinks : Hyperlinks.Concat(hyperlinks));
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString()
            => $"Document: (Charts: {Charts.Count()}, Images: {Images.Count()}, Hyperlinks: {Hyperlinks.Count()})";

        /// <summary>
        ///
        /// </summary>
        /// <param name="archive">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] ZipArchive archive)
        {
            if (archive is null)
                throw new ArgumentNullException(nameof(archive));

            using (Stream stream = archive.GetEntry("word/document.xml")?.Open())
            {
                Content.Save(stream);
            }

            foreach (ChartInfo item in Charts)
            {
                item.Save(archive);
            }

            foreach (ImageInfo item in Images)
            {
                item.Save(archive);
            }
        }
    }
}