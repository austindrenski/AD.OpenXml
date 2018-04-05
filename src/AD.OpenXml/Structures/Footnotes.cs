using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

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
        private static readonly StringSegment MimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";

        /// <summary>
        ///
        /// </summary>
        private static readonly StringSegment SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";

        /// <summary>
        ///
        /// </summary>
        public readonly StringSegment RelationId;

        /// <summary>
        ///
        /// </summary>
        public readonly StringSegment Target = "footnotes.xml";

        /// <summary>
        /// The XML file located at: /word/footnotes.xml.
        /// </summary>
        [NotNull]
        public XElement Content { get; }

        /// <summary>
        /// The hyperlinks listed in: /word/_rels/footnotes.xml.rels.
        /// </summary>
        [NotNull]
        public IImmutableSet<HyperlinkInfo> Hyperlinks { get; }

        /// <summary>
        ///
        /// </summary>
        public StringSegment PartName => $"/word/{Target}";

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
        /// <param name="rId">
        ///
        /// </param>
        /// <param name="content">
        ///
        /// </param>
        /// <param name="hyperlinks">
        ///
        /// </param>
        public Footnotes(StringSegment rId, [NotNull] XElement content, [NotNull] IEnumerable<HyperlinkInfo> hyperlinks)
        {
            if (content is null)
            {
                throw new ArgumentNullException(nameof(content));
            }

            if (hyperlinks is null)
            {
                throw new ArgumentNullException(nameof(hyperlinks));
            }

            RelationId = rId;
            Content = content;
            Hyperlinks = hyperlinks.ToImmutableHashSet();
        }

        /// <summary>
        /// Initializes an <see cref="OpenXmlPackageVisitor"/> by reading Footnotes parts into memory.
        /// </summary>
        /// <param name="archive">
        /// The archive to which changes can be saved.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public Footnotes([NotNull] ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            RelationId =
                archive.ReadXml(DocumentRelsInfo.Path)
                       .Elements(DocumentRelsInfo.Elements.Relationship)
                       .Single(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Type) == SchemaType)
                       .Attribute(DocumentRelsInfo.Attributes.Id)
                       .Value;

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
                       .ToImmutableHashSet();
        }

        ///  <summary>
        ///
        ///  </summary>
        /// <param name="content"></param>
        ///  <param name="hyperlinks"></param>
        ///  <returns></returns>
        public Footnotes With([CanBeNull] XElement content = default, [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
        {
            return new Footnotes(RelationId, content ?? Content, hyperlinks ?? Hyperlinks);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        public Footnotes Concat([NotNull] Footnotes other)
        {
            return Concat(other.Content, other.Hyperlinks);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="hyperlinks"></param>
        /// <returns></returns>
        public Footnotes Concat([CanBeNull] XElement content = default, [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
        {
            XElement footnotes =
                content is default
                    ? Content
                    : new XElement(
                        Content.Name,
                        Content.Attributes(),
                        Content.Elements(),
                        content.Elements());

            return new Footnotes(RelationId, footnotes, hyperlinks is default ? Hyperlinks : Hyperlinks.Concat(hyperlinks));
        }

//        /// <inheritdoc />
//        [Pure]
//        [NotNull]
//        public override string ToString()
//        {
//            return ToXElement().ToString();
//        }

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
            {
                throw new ArgumentNullException(nameof(archive));
            }

            using (Stream stream = archive.GetEntry(PartName.Substring(1)).Open())
            {
                Content.Save(stream);
            }
        }
    }
}