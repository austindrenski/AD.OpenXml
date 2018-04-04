using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
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
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public XElement Content { get; }

        /// <summary>
        /// word/charts/chart#.xml.
        /// </summary>
        [NotNull]
        public IImmutableSet<ChartInfo> Charts { get; }

        /// <summary>
        /// word/media/image#.[jpeg|png|svg].
        /// </summary>
        [NotNull]
        public IImmutableSet<ImageInfo> Images { get; }

        /// <summary>
        /// Hyperlinks
        /// </summary>
        [NotNull]
        public IImmutableSet<HyperlinkInfo> Hyperlinks { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public Relationships Relationships =>
            new Relationships(
                Charts.Select(x => x.RelationshipEntry),
                Images.Select(x => x.RelationshipEntry),
                Hyperlinks.Select(x => x.RelationshipEntry));

        /// <summary>
        ///
        /// </summary>
        /// <param name="content">
        ///
        /// </param>
        /// <param name="charts">
        ///
        /// </param>
        /// <param name="images">
        ///
        /// </param>
        /// <param name="hyperlinks">
        ///
        /// </param>
        public Document([NotNull] XElement content, [NotNull] IEnumerable<ChartInfo> charts, [NotNull] IEnumerable<ImageInfo> images, [NotNull] IEnumerable<HyperlinkInfo> hyperlinks)
        {
            if (content is null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            if (charts is null)
            {
                throw new ArgumentNullException(nameof(charts));
            }

            if (images is null)
            {
                throw new ArgumentNullException(nameof(images));
            }

            if (hyperlinks is null)
            {
                throw new ArgumentNullException(nameof(hyperlinks));
            }

            Content = content;
            Charts = charts.ToImmutableHashSet();
            Images = images.ToImmutableHashSet();
            Hyperlinks = hyperlinks.ToImmutableHashSet();
        }

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
            {
                throw new ArgumentNullException(nameof(archive));
            }

            Content = archive.ReadXml();

            XElement documentRelations = archive.ReadXml(DocumentRelsInfo.Path);

            Charts =
                documentRelations.Elements()
                                 .Select(
                                     x =>
                                         new
                                         {
                                             Id = (string) x.Attribute(DocumentRelsInfo.Attributes.Id),
                                             Target = (string) x.Attribute(DocumentRelsInfo.Attributes.Target)
                                         })
                                 .Where(x => x.Target.StartsWith("charts/"))
                                 .Select(x => new ChartInfo(x.Id, archive.ReadXml($"word/{x.Target}")))
                                 .ToImmutableHashSet();

            Images =
                documentRelations.Elements()
                                 .Select(
                                     x =>
                                         new
                                         {
                                             Id = (string) x.Attribute(DocumentRelsInfo.Attributes.Id),
                                             Target = (string) x.Attribute(DocumentRelsInfo.Attributes.Target)
                                         })
                                 .Where(x => x.Target.StartsWith("media/"))
                                 .Select(x => ImageInfo.Create(x.Id, x.Target, archive.ReadByteArray($"word/{x.Target}")))
                                 .ToImmutableHashSet();

            Hyperlinks =
                documentRelations.Elements()
                                 .Select(
                                     x =>
                                         new
                                         {
                                             Id = (string) x.Attribute(DocumentRelsInfo.Attributes.Id),
                                             Target = (string) x.Attribute(DocumentRelsInfo.Attributes.Target),
                                             TargetMode = (string) x.Attribute("TargetMode")
                                         })
                                 .Where(x => x.TargetMode != null)
                                 .Select(x => HyperlinkInfo.Create(x.Id, x.Target, x.TargetMode))
                                 .ToImmutableHashSet();
        }

//        /// <inheritdoc />
//        [Pure]
//        [NotNull]
//        public override string ToString()
//        {
//            return ToXElement().ToString();
//        }
    }
}