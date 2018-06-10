using System;
using System.Collections.Generic;
using System.IO;
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
    public class Footnotes
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
        ///
        /// </summary>
        [NotNull] private static readonly string MimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] private static readonly string SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public readonly string RelationId;

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly string Target = "footnotes.xml";

        /// <summary>
        /// The XML file located at: /word/footnotes.xml.
        /// </summary>
        [NotNull]
        public XElement Content { get; }

        /// <summary>
        /// The hyperlinks listed in: /word/_rels/footnotes.xml.rels.
        /// </summary>
        [NotNull]
        public IEnumerable<HyperlinkInfo> Hyperlinks { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string PartName => $"/word/{Target}";

        /// <summary>
        ///
        /// </summary>
        public ContentTypes.Override ContentTypeEntry => new ContentTypes.Override(PartName, MimeType);

        /// <summary>
        ///
        /// </summary>
        public Relationships.Entry RelationshipEntry => new Relationships.Entry(RelationId, Target, SchemaType);

        /// <summary>
        /// The current footnote count.
        /// </summary>
        public int Count =>
            Content.Elements()
                   .SkipWhile(x => (int) x.Attribute(W + "id") <= 0)
                   .Count();

        /// <summary>
        /// The number of relationships in the Footnotes.
        /// </summary>
        public int RelationshipsMax =>
            Hyperlinks.Select(x => x.NumericId)
                      .DefaultIfEmpty(default)
                      .Max();

        /// <summary>
        /// The current revision number.
        /// </summary>
        public int RevisionId =>
            Content.Descendants()
                   .Where(x => Revisions.Contains(x.Name))
                   .Select(x => (int) x.Attribute(W + "id"))
                   .DefaultIfEmpty(default)
                   .Max();

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <param name="content"></param>
        /// <param name="hyperlinks"></param>
        public Footnotes([NotNull] string rId, [NotNull] XElement content, [NotNull] IEnumerable<HyperlinkInfo> hyperlinks)
        {
            if (rId is null)
                throw new ArgumentNullException(nameof(rId));

            if (content is null)
                throw new ArgumentNullException(nameof(content));

            if (hyperlinks is null)
                throw new ArgumentNullException(nameof(hyperlinks));

            RelationId = rId;
            Content = content;
            Hyperlinks = hyperlinks.ToArray();
        }

        /// <summary>
        /// Initializes an <see cref="OpenXmlPackageVisitor"/> by reading Footnotes parts into memory.
        /// </summary>
        /// <param name="archive">The archive to which changes can be saved.</param>
        /// <exception cref="ArgumentNullException"/>
        public Footnotes([NotNull] ZipArchive archive)
        {
            if (archive is null)
                throw new ArgumentNullException(nameof(archive));

            RelationId =
                archive.ReadXml(DocumentRelsInfo.Path, Relationships.Empty.ToXElement())
                       .Elements(DocumentRelsInfo.Elements.Relationship)
                       .SingleOrDefault(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Type) == SchemaType)?
                       .Attribute(DocumentRelsInfo.Attributes.Id)?
                       .Value ?? string.Empty;

            // TODO: hard-coding to rId1 until other package parts are migrated.
            RelationId = "rId1";

            Content = archive.ReadXml(PartName.Substring(1), new XElement(W + "footnotes"));

            Hyperlinks =
                archive.ReadXml(FootnotesRelsInfo.Path, Relationships.Empty.ToXElement())
                       .Elements()
                       .Select(
                           x =>
                               new
                               {
                                   Id = (string) x.Attribute(FootnotesRelsInfo.Attributes.Id),
                                   Target = (string) x.Attribute(FootnotesRelsInfo.Attributes.Target),
                                   TargetMode = (string) x.Attribute(FootnotesRelsInfo.Attributes.TargetMode)
                               })
                       .Where(x => x.TargetMode != null)
                       .Select(x => new HyperlinkInfo(x.Id, x.Target, x.TargetMode))
                       .ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="hyperlinks"></param>
        /// <returns>
        ///
        /// </returns>
        public Footnotes With([CanBeNull] XElement content = default, [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
            => new Footnotes(RelationId, content ?? Content, hyperlinks ?? Hyperlinks);

        /// <summary>
        ///
        /// </summary>
        /// <param name="other"></param>
        /// <returns>
        ///
        /// </returns>
        public Footnotes Concat([NotNull] Footnotes other) => Concat(other.Content, other.Hyperlinks);

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="hyperlinks"></param>
        /// <returns>
        ///
        /// </returns>
        public Footnotes Concat([CanBeNull] XElement content = default, [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
            => new Footnotes(
                RelationId,
                content is null
                    ? Content
                    : new XElement(
                        Content.Name,
                        Content.Attributes(),
                        Content.Elements(),
                        content.Elements()),
                hyperlinks is null
                    ? Hyperlinks
                    : Hyperlinks.Concat(hyperlinks));

        /// <summary>
        ///
        /// </summary>
        /// <param name="archive"></param>
        /// <exception cref="ArgumentNullException" />
        /// <exception cref="FileNotFoundException" />
        public void Save([NotNull] ZipArchive archive)
        {
            if (archive is null)
                throw new ArgumentNullException(nameof(archive));

            using (Stream stream =
                archive.GetEntry(PartName.Substring(1))?.Open() ??
                archive.CreateEntry(PartName.Substring(1)).Open())
            {
                Content.Save(stream);
            }
        }
    }
}