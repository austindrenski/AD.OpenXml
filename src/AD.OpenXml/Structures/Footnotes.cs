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
        [NotNull] public const string ContentType =
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
        /// The package that initialized the <see cref="Footnotes"/>.
        /// </summary>
        [NotNull] private readonly Package _package;

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
        /// The current revision number.
        /// </summary>
        public int RevisionId =>
            Content.Descendants()
                   .Where(x => Revisions.Contains(x.Name))
                   .Select(x => (int) x.Attribute(W + "id"))
                   .DefaultIfEmpty(default)
                   .Max();

        /// <summary>
        /// Initializes an <see cref="OpenXmlPackageVisitor"/> by reading Footnotes parts into memory.
        /// </summary>
        /// <param name="package">The package to which changes can be saved.</param>
        /// <exception cref="ArgumentNullException"/>
        public Footnotes([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            _package = package;

            if (package.PartExists(PartName))
            {
                using (Stream stream = package.GetPart(PartName).GetStream())
                {
                    Content = XElement.Load(stream);
                }
            }
            else
            {
                Content =
                    new XElement(W + "footnotes",
                        new XAttribute(XNamespace.Xmlns + "w", W));
            }

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
        /// <param name="package"></param>
        /// <param name="content"></param>
        /// <param name="hyperlinks"></param>
        private Footnotes(
            [NotNull] Package package,
            [NotNull] XElement content,
            [NotNull] IEnumerable<HyperlinkInfo> hyperlinks)
        {
            _package = package;
            Content = content;
            Hyperlinks = hyperlinks.ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="hyperlinks"></param>
        /// <returns>
        ///
        /// </returns>
        public Footnotes With(
            [CanBeNull] XElement content = default,
            [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
            => new Footnotes(
                _package,
                content ?? Content,
                hyperlinks ?? Hyperlinks);

        /// <summary>
        ///
        /// </summary>
        /// <param name="other"></param>
        /// <returns>
        ///
        /// </returns>
        public Footnotes Concat([NotNull] Footnotes other)
        {
            if (other is null)
                throw new ArgumentNullException(nameof(other));

            Package result = Package.Open(_package.ToStream(), FileMode.Open);

            PackagePart footnotesPart =
                result.PartExists(PartName)
                    ? result.GetPart(PartName)
                    : result.CreatePart(PartName, ContentType);

            Dictionary<string, PackageRelationship> resources = new Dictionary<string, PackageRelationship>();

            foreach (HyperlinkInfo info in other.Hyperlinks)
            {
                resources[info.Id] = footnotesPart.CreateRelationship(info.Target, info.TargetMode, HyperlinkInfo.RelationshipType);
            }

            XElement content =
                new XElement(
                    Content.Name,
                    Combine(Content.Attributes(), other.Content.Attributes()),
                    Content.Nodes(),
                    other.Content.Nodes().Select(x => Update(x, resources)));

            using (Stream stream = footnotesPart.GetStream())
            {
                using (XmlWriter xml = XmlWriter.Create(stream, XmlWriterSettings))
                {
                    content.WriteTo(xml);
                }
            }

            return new Footnotes(result);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="target">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException" />
        public void CopyTo([NotNull] Package target)
        {
            if (target is null)
                throw new ArgumentNullException(nameof(target));

            if (target.PartExists(PartName))
                target.DeletePart(PartName);

            target.GetPart(Document.PartName)
                  .CreateRelationship(PartName, TargetMode.Internal, RelationshipType);

            PackagePart source = _package.GetPart(PartName);
            PackagePart destination = target.CreatePart(PartName, ContentType);


            using (Stream from = source.GetStream())
            {
                using (Stream to = destination.GetStream())
                {
                    from.CopyTo(to);
                }
            }

            foreach (PackageRelationship relationship in source.GetRelationships())
            {
                destination.CreateRelationship(
                    relationship.TargetUri,
                    relationship.TargetMode,
                    relationship.RelationshipType,
                    relationship.Id);

                if (relationship.RelationshipType == HyperlinkInfo.RelationshipType)
                    continue;

                PackagePart part = _package.GetPart(PartUri(relationship.TargetUri));

                using (Stream from = part.GetStream())
                {
                    using (Stream to = target.CreatePart(part.Uri, part.ContentType).GetStream())
                    {
                        from.CopyTo(to);
                    }
                }
            }
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString() => $"(Footnotes: {Count}, Hyperlinks: {Hyperlinks.Count()})";

        /// <summary>
        /// Constructs the part URI from the target URI.
        /// </summary>
        /// <param name="targetUri">The relative target URI.</param>
        /// <returns>
        /// The relative part URI.
        /// </returns>
        [Pure]
        [NotNull]
        private static Uri PartUri([NotNull] Uri targetUri) => new Uri($"/word/{targetUri}", UriKind.Relative);

        [Pure]
        [NotNull]
        private static XAttribute[] Combine([NotNull] IEnumerable<XAttribute> source, [NotNull] IEnumerable<XAttribute> others)
        {
            XAttribute[] attributes = source as XAttribute[] ?? source.ToArray();

            IEnumerable<XAttribute> otherAttributes =
                others.Where(x => !attributes.Any(y => y.Name == x.Name || y.Value == x.Value));

            return attributes.Concat(otherAttributes).ToArray();
        }

        [Pure]
        [NotNull]
        private static XObject Update(XObject xObject, Dictionary<string, PackageRelationship> resources)
        {
            switch (xObject)
            {
                case XAttribute a
                    when (a.Name == R + "id" || a.Name == R + "embed") &&
                         resources.TryGetValue(a.Value, out PackageRelationship rel):
                    return new XAttribute(a.Name, rel.Id);

                case XElement e:
                    return
                        new XElement(
                            e.Name,
                            e.Attributes().Select(x => Update(x, resources)),
                            e.Nodes().Select(x => Update(x, resources)));

                default:
                    return xObject;
            }
        }
    }
}