using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
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
        [NotNull] public static readonly Uri PartUri = new Uri("/word/footnotes.xml", UriKind.Relative);

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

            if (!package.FileOpenAccess.HasFlag(FileAccess.Read))
                throw new IOException("The package is write-only.");

            _package =
                package.FileOpenAccess.HasFlag(FileAccess.Write)
                    ? package.ToPackage()
                    : package;

            if (package.PartExists(PartUri))
            {
                using (Stream stream = package.GetPart(PartUri).GetStream())
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
                package.PartExists(PartUri)
                    ? package.GetPart(PartUri)
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
        [Pure]
        [NotNull]
        public Footnotes With([CanBeNull] XElement content = default, [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
        {
            Package package = _package.ToPackage(FileAccess.ReadWrite);

            if (package.PartExists(PartUri))
                package.DeletePart(PartUri);

            PackagePart footnotesPart = package.CreatePart(PartUri, ContentType);

            foreach (HyperlinkInfo info in hyperlinks ?? Hyperlinks)
            {
                footnotesPart.CreateRelationship(info.Target, info.TargetMode, HyperlinkInfo.RelationshipType, info.Id);
            }

            (content ?? Content).WriteTo(footnotesPart);

            return new Footnotes(package);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="other"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public Footnotes Concat([NotNull] Footnotes other) => Concat(this, other);

        /// <summary>
        ///
        /// </summary>
        /// <param name="first"></param>
        /// <param name="second"></param>
        /// <returns>
        ///
        /// </returns>
        public static Footnotes Concat([NotNull] Footnotes first, [NotNull] Footnotes second)
        {
            if (first is null)
                throw new ArgumentNullException(nameof(first));
            if (second is null)
                throw new ArgumentNullException(nameof(second));

            Sequence idSequence = new Sequence("rId{0}");
            MappingSequence footnoteSequence = new MappingSequence();

            Package result = first._package.ToPackage(FileAccess.ReadWrite);

            if (result.PartExists(PartUri))
                result.DeletePart(PartUri);

            PackagePart part = result.CreatePart(PartUri, ContentType);

            Dictionary<string, PackageRelationship> resources = new Dictionary<string, PackageRelationship>();
            Dictionary<string, PackageRelationship> otherResources = new Dictionary<string, PackageRelationship>();

            // For existing hyperlinks: recreate the relationship.
            foreach (HyperlinkInfo info in first.Hyperlinks)
            {
                resources[info.Id] =
                    part.CreateRelationship(info.Target, info.TargetMode, HyperlinkInfo.RelationshipType, idSequence.NextValue());
            }

            // For new hyperlinks: create the relationship.
            foreach (HyperlinkInfo info in second.Hyperlinks)
            {
                otherResources[info.Id] =
                    part.CreateRelationship(info.Target, info.TargetMode, HyperlinkInfo.RelationshipType, idSequence.NextValue());
            }

            XElement content =
                new XElement(
                    first.Content.Name,
                    Combine(first.Content.Attributes(), second.Content.Attributes()),
                    first.Content.Nodes()
                         .Select(x => UpdateResources(x, resources))
                         .Select(x => UpdateFootnotes(x, footnoteSequence)),
                    second.Content
                          .Nodes()
                          .Select(x => UpdateResources(x, otherResources))
                          .Select(x => UpdateFootnotes(x, footnoteSequence)));

            content.WriteTo(part);

            PackagePart documentPart = result.GetPart(Document.PartUri);

            XElement document;

            using (Stream stream = documentPart.GetStream())
            {
                document = XElement.Load(stream);
            }

            XElement updatedDocument =
                new XElement(
                    document.Name,
                    document.Attributes(),
                    document.Nodes().Select(x => UpdateReferences(x, footnoteSequence)));

            updatedDocument.WriteTo(documentPart);

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

            if (!target.FileOpenAccess.HasFlag(FileAccess.Read))
                throw new IOException("The package is write-only.");

            if (!_package.PartExists(PartUri))
                return;

            if (target.PartExists(PartUri))
                target.DeletePart(PartUri);

            target.GetPart(Document.PartUri)
                  .CreateRelationship(PartUri, TargetMode.Internal, RelationshipType);

            PackagePart source = _package.GetPart(PartUri);
            PackagePart destination = target.CreatePart(PartUri, ContentType);

            using (Stream from = source.GetStream())
            {
                using (Stream to = destination.GetStream())
                {
                    from.CopyTo(to);
                }
            }

            foreach (PackageRelationship relationship in source.GetRelationships())
            {
                if (relationship.RelationshipType != ChartInfo.RelationshipType &&
                    relationship.RelationshipType != ImageInfo.RelationshipType &&
                    relationship.RelationshipType != HyperlinkInfo.RelationshipType)
                    continue;

                destination.CreateRelationship(
                    relationship.TargetUri,
                    relationship.TargetMode,
                    relationship.RelationshipType,
                    relationship.Id);

                if (relationship.RelationshipType != ChartInfo.RelationshipType &&
                    relationship.RelationshipType != ImageInfo.RelationshipType)
                    continue;

                PackagePart part = _package.GetPart(MakePartUri(relationship.TargetUri));

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
        public override string ToString() => $"(Footnotes: {Content.Elements().Count()}, Hyperlinks: {Hyperlinks.Count()})";

        /// <summary>
        /// Constructs the part URI from the target URI.
        /// </summary>
        /// <param name="targetUri">The relative target URI.</param>
        /// <returns>
        /// The relative part URI.
        /// </returns>
        [Pure]
        [NotNull]
        private static Uri MakePartUri([NotNull] Uri targetUri) => new Uri($"/word/{targetUri}", UriKind.Relative);

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
        private static XObject UpdateResources(XObject xObject, Dictionary<string, PackageRelationship> resources)
        {
            switch (xObject)
            {
                case XAttribute a when a.Name == R + "id" || a.Name == R + "embed":
                    return new XAttribute(a.Name, resources.TryGetValue(a.Value, out PackageRelationship rel) ? rel.Id : a.Value);

                case XAttribute a:
                    return new XAttribute(a.Name, a.Value);

                case XElement e:
                    return
                        new XElement(
                            e.Name,
                            e.Attributes().Select(x => UpdateResources(x, resources)),
                            e.Nodes().Select(x => UpdateResources(x, resources)));

                case XText t:
                    return new XText(t.Value);

                default:
                    return xObject;
            }
        }

        [Pure]
        [NotNull]
        private static XObject UpdateFootnotes([NotNull] XObject xObject, MappingSequence footnoteSequence)
        {
            switch (xObject)
            {
                case XAttribute a when a.Name == W + "id" && a.Parent?.Name == W + "footnote":
                    return new XAttribute(a.Name, footnoteSequence.NextValue(a.Value));

                case XAttribute a:
                    return new XAttribute(a.Name, a.Value);

                case XElement e:
                    return
                        new XElement(
                            e.Name,
                            e.Attributes().Select(x => UpdateFootnotes(x, footnoteSequence)),
                            e.Nodes().Select(x => UpdateFootnotes(x, footnoteSequence)));

                case XText t:
                    return new XText(t.Value);

                default:
                    return xObject;
            }
        }

        [Pure]
        [NotNull]
        private static XObject UpdateReferences([NotNull] XObject xObject, MappingSequence footnoteSequence)
        {
            switch (xObject)
            {
                case XAttribute a when a.Name == W + "id" && a.Parent?.Name == W + "footnoteReference":
                    return new XAttribute(a.Name, footnoteSequence.GetValue(a.Value));

                case XAttribute a:
                    return new XAttribute(a.Name, a.Value);

                case XElement e:
                    return
                        new XElement(
                            e.Name,
                            e.Attributes().Select(x => UpdateReferences(x, footnoteSequence)),
                            e.Nodes().Select(x => UpdateReferences(x, footnoteSequence)));

                case XText t:
                    return new XText(t.Value);

                default:
                    return xObject;
            }
        }
    }
}