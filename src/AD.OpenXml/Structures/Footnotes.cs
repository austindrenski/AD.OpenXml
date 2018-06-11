using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
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
        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] private static readonly XmlWriterSettings XmlWriterSettings =
            new XmlWriterSettings
            {
                Async = false,
                DoNotEscapeUriAttributes = false,
                CheckCharacters = true,
                CloseOutput = false,
                ConformanceLevel = ConformanceLevel.Document,
                Encoding = Encoding.UTF8,
                Indent = false,
                IndentChars = "  ",
                NamespaceHandling = NamespaceHandling.OmitDuplicates,
                NewLineChars = Environment.NewLine,
                NewLineHandling = NewLineHandling.None,
                NewLineOnAttributes = false,
                OmitXmlDeclaration = false,
                WriteEndDocumentOnClose = true
            };

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
        [NotNull] public const string MimeType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string RelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly Uri PartName = new Uri("/word/footnotes.xml", UriKind.Relative);

        /// <summary>
        ///
        /// </summary>
        [NotNull] public readonly string RelationId;

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
        /// <param name="package">The package to which changes can be saved.</param>
        /// <exception cref="ArgumentNullException"/>
        public Footnotes([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            RelationId =
                package.PartExists(DocumentRelsInfo.PartName)
                    ? XElement.Load(package.GetPart(DocumentRelsInfo.PartName).GetStream())
                              .Elements(DocumentRelsInfo.Elements.Relationship)
                              .SingleOrDefault(x => (string) x.Attribute(DocumentRelsInfo.Attributes.Type) == RelationshipType)?
                              .Attribute(DocumentRelsInfo.Attributes.Id)?
                              .Value ?? string.Empty
                    : string.Empty;

            // TODO: hard-coding to rId1 until other package parts are migrated.
            RelationId = "rId1";

            Content =
                package.PartExists(PartName)
                    ? XElement.Load(package.GetPart(PartName).GetStream())
                    : new XElement(W + "footnotes",
                        new XAttribute(XNamespace.Xmlns + "w", W));

            Hyperlinks =
                package.PartExists(PartName)
                    ? package.GetPart(PartName)
                             .GetRelationshipsByType(HyperlinkInfo.RelationshipType)
                             .Select(x => new HyperlinkInfo(x.Id, x.TargetUri, x.TargetMode))
                             .ToArray()
                    : new HyperlinkInfo[0];
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
        public Footnotes Concat([NotNull] Footnotes other) =>
            new Footnotes(
                RelationId,
                new XElement(
                    Content.Name,
                    Content.Attributes(),
                    Content.Nodes(),
                    other.Content.Nodes()),
                Hyperlinks.Concat(other.Hyperlinks));

        /// <summary>
        ///
        /// </summary>
        /// <param name="package"></param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            PackagePart document = package.GetPart(Document.PartName);

            if (!document.RelationshipExists(RelationId))
                document.CreateRelationship(PartName, TargetMode.Internal, RelationshipType);

            PackagePart footnotes =
                package.PartExists(PartName)
                    ? package.GetPart(PartName)
                    : package.CreatePart(PartName, MimeType);

            (string from, string to)[] updates =
                Hyperlinks
                    .Select(
                        x =>
                            (from: x.RelationId,
                                to: footnotes.RelationshipExists(x.RelationId) &&
                                    footnotes.GetRelationship(x.RelationId).TargetUri == x.Target
                                    ? footnotes.GetRelationship(x.RelationId).Id
                                    : footnotes.CreateRelationship(x.Target, x.TargetMode, HyperlinkInfo.RelationshipType).Id))
                    .Select(x => (x.from, x.to))
                    .ToArray();

            XAttribute[] attributes =
                Content.DescendantsAndSelf()
                       .Attributes()
                       .Where(x => x.Name == R + "id")
                       .ToArray();
            for (int i = 0;
                 i < attributes.Length;
                 i++)
            {
                XAttribute attribute = attributes[i];
                foreach ((string from, string to) item in updates)
                {
                    if (attribute.Value != item.from)
                        continue;

                    attribute.SetValue(item.to);
                    break;
                }
            }

            using (Stream stream = footnotes.GetStream())
            {
                using (XmlWriter xml = XmlWriter.Create(stream, XmlWriterSettings))
                {
                    Content.WriteTo(xml);
                }
            }
        }
    }
}