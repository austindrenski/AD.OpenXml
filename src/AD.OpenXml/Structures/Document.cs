﻿using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
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
        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;
        [NotNull] private static readonly XNamespace WP = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

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
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly Uri PartUri = new Uri("/word/document.xml", UriKind.Relative);

        /// <summary>
        /// The package that initialized the <see cref="Document"/>.
        /// </summary>
        [NotNull] private readonly Package _package;

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
        public IDictionary<string, XElement> ChartReferences => Charts.ToDictionary(x => x.Id, x => x.Chart);

        /// <summary>
        /// Maps image reference id to image node.
        /// </summary>
        [NotNull]
        public IDictionary<string, (string contentType, string description, string base64)> ImageReferences =>
            Images.ToDictionary(x => x.Id.ToString(), x => (x.ContentType, string.Empty, x.ToBase64String()));

        /// <summary>
        /// Initializes an <see cref="OpenXmlPackageVisitor"/> by reading document parts into memory.
        /// </summary>
        /// <param name="package">
        /// The package to which changes can be saved.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public Document([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            if (!package.FileOpenAccess.HasFlag(FileAccess.Read))
                throw new IOException("The package is write-only.");

            _package =
                package.FileOpenAccess.HasFlag(FileAccess.Write)
                    ? package.ToPackage()
                    : package;

            if (_package.PartExists(PartUri))
                Content = _package.GetPart(PartUri).ReadXml();
            else
            {
                Content =
                    new XElement(W + "document",
                        new XAttribute(XNamespace.Xmlns + "w", W),
                        new XElement(W + "body"));
            }

            Charts =
                _package.GetPart(PartUri)
                        .GetRelationshipsByType(ChartInfo.RelationshipType)
                        .Select(x => ChartInfo.Read(_package.GetPart(MakePartUri(x.TargetUri)), x))
                        .ToArray();

            Images =
                _package.GetPart(PartUri)
                        .GetRelationshipsByType(ImageInfo.RelationshipType)
                        .Select(x => ImageInfo.Read(_package.GetPart(MakePartUri(x.TargetUri)), x))
                        .ToArray();

            Hyperlinks =
                _package.GetPart(PartUri)
                        .GetRelationshipsByType(HyperlinkInfo.RelationshipType)
                        .Select(x => new HyperlinkInfo(x.Id, x.TargetUri, x.TargetMode))
                        .ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="content"></param>
        /// <param name="charts"></param>
        /// <param name="images"></param>
        /// <param name="hyperlinks"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public Document With(
            [CanBeNull] XElement content = default,
            [CanBeNull] IEnumerable<ChartInfo> charts = default,
            [CanBeNull] IEnumerable<ImageInfo> images = default,
            [CanBeNull] IEnumerable<HyperlinkInfo> hyperlinks = default)
        {
            Package package = _package.ToPackage(FileAccess.ReadWrite);

            if (package.PartExists(PartUri))
                package.DeletePart(PartUri);

            PackagePart documentPart = package.CreatePart(PartUri, ContentType);

            foreach (ChartInfo info in charts ?? Charts)
            {
                if (package.PartExists(MakePartUri(info.TargetUri)))
                    package.DeletePart(MakePartUri(info.TargetUri));

                documentPart.CreateRelationship(info.TargetUri, TargetMode.Internal, ChartInfo.RelationshipType, info.Id);

                info.Chart.WriteTo(package.CreatePart(MakePartUri(info.TargetUri), ChartInfo.ContentType));
            }

            foreach (ImageInfo info in images ?? Images)
            {
                if (package.PartExists(MakePartUri(info.TargetUri)))
                    package.DeletePart(MakePartUri(info.TargetUri));

                documentPart.CreateRelationship(info.TargetUri, TargetMode.Internal, ImageInfo.RelationshipType, info.Id);

                info.WriteTo(package.CreatePart(MakePartUri(info.TargetUri), info.ContentType));
            }

            foreach (HyperlinkInfo info in hyperlinks ?? Hyperlinks)
            {
                documentPart.CreateRelationship(info.Target, info.TargetMode, HyperlinkInfo.RelationshipType, info.Id);
            }

            (content ?? Content).WriteTo(documentPart);

            return new Document(package);
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
        public Document Concat([NotNull] Document other) => Concat(this, other);

        /// <summary>
        ///
        /// </summary>
        /// <param name="first"></param>
        /// <param name="second"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public static Document Concat([NotNull] Document first, [NotNull] Document second)
        {
            if (second is null)
                throw new ArgumentNullException(nameof(second));

            Package result = first._package.ToPackage(FileAccess.ReadWrite);

            PackagePart part = result.GetPart(PartUri);

            Sequence idSequence = new Sequence("rId{0}");
            Sequence chartSequence = new Sequence();
            Sequence imageSequence = new Sequence();

            PackageRelationship[] allRelationships =
                part.GetRelationships()
                    .ToArray();

            // Remove all part relationships.
            foreach (PackageRelationship rel in allRelationships)
            {
                part.DeleteRelationship(rel.Id);
            }

            Dictionary<string, PackageRelationship> resources = new Dictionary<string, PackageRelationship>();
            Dictionary<string, PackageRelationship> otherResources = new Dictionary<string, PackageRelationship>();

            // Re-number relationships we don't modify.
            PackageRelationship[] defaultRelationships =
                allRelationships.Where(x => x.RelationshipType != ChartInfo.RelationshipType)
                                .Where(x => x.RelationshipType != ImageInfo.RelationshipType)
                                .Where(x => x.RelationshipType != HyperlinkInfo.RelationshipType)
                                .ToArray();

            foreach (PackageRelationship rel in defaultRelationships)
            {
                resources[rel.Id] =
                    part.CreateRelationship(rel.TargetUri, rel.TargetMode, rel.RelationshipType, idSequence.NextValue());
            }

            // For existing charts: remove, replace, and recreate the relationship.
            foreach (ChartInfo info in first.Charts)
            {
                result.DeletePart(MakePartUri(info.TargetUri));

                Uri uri = info.MakeUri(chartSequence.NextValue());

                info.Chart.WriteTo(result.CreatePart(MakePartUri(uri), ChartInfo.ContentType));

                resources[info.Id] =
                    part.CreateRelationship(uri, TargetMode.Internal, ChartInfo.RelationshipType, idSequence.NextValue());
            }

            // For new charts: add then create the relationship.
            foreach (ChartInfo info in second.Charts)
            {
                Uri uri = info.MakeUri(chartSequence.NextValue());

                info.Chart.WriteTo(result.CreatePart(MakePartUri(uri), ChartInfo.ContentType));

                otherResources[info.Id] =
                    part.CreateRelationship(uri, TargetMode.Internal, ChartInfo.RelationshipType, idSequence.NextValue());
            }

            // For existing images: remove, replace, and recreate the relationship.
            foreach (ImageInfo info in first.Images)
            {
                result.DeletePart(MakePartUri(info.TargetUri));

                Uri uri = info.MakeUri(imageSequence.NextValue());

                info.WriteTo(result.CreatePart(MakePartUri(uri), info.ContentType));

                resources[info.Id] =
                    part.CreateRelationship(uri, TargetMode.Internal, ImageInfo.RelationshipType, idSequence.NextValue());
            }

            // For new images: add then create the relationship.
            foreach (ImageInfo info in second.Images)
            {
                Uri uri = info.MakeUri(imageSequence.NextValue());

                info.WriteTo(result.CreatePart(MakePartUri(uri), info.ContentType));

                otherResources[info.Id] =
                    part.CreateRelationship(uri, TargetMode.Internal, ImageInfo.RelationshipType, idSequence.NextValue());
            }

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
                    Update(
                        new XElement(
                            W + "body",
                            first.Content.Element(W + "body")?.Elements().Select(x => Update(x, resources)),
                            second.Content.Element(W + "body")?.Elements().Select(x => Update(x, otherResources))),
                        new Sequence()));

            content.RemoveDuplicateSectionProperties();

            content.WriteTo(part);

            return new Document(result);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="target">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException" />
        /// <exception cref="IOException" />
        public void CopyTo([NotNull] Package target)
        {
            if (target is null)
                throw new ArgumentNullException(nameof(target));

            if (!target.FileOpenAccess.HasFlag(FileAccess.Write))
                throw new IOException("The package is read-only.");

            if (target.PartExists(PartUri))
                target.DeletePart(PartUri);

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
        [NotNull]
        public override string ToString()
            => $"(Charts: {Charts.Count()}, Images: {Images.Count()}, Hyperlinks: {Hyperlinks.Count()})";

        [Pure]
        [NotNull]
        static XObject Update([NotNull] XObject xObject, [NotNull] Dictionary<string, PackageRelationship> resources)
        {
            switch (xObject)
            {
                case XAttribute a
                    when (a.Name == R + "id" || a.Name == R + "embed") &&
                         resources.TryGetValue(a.Value, out PackageRelationship rel):
                    return new XAttribute(a.Name, rel.Id);

                case XAttribute a:
                    return new XAttribute(a.Name, a.Value);

                case XElement e:
                    return
                        new XElement(
                            e.Name,
                            e.Attributes().Select(x => Update(x, resources)),
                            e.Nodes().Select(x => Update(x, resources)));

                case XText t:
                    return new XText(t.Value);

                default:
                    return xObject;
            }
        }

        [Pure]
        [NotNull]
        static XObject Update([NotNull] XObject xObject, Sequence docPrSequence)
        {
            switch (xObject)
            {
                case XAttribute a
                    when a.Name == "id" && a.Parent?.Name == WP + "docPr":
                    return new XAttribute(a.Name, docPrSequence.NextValue());

                case XAttribute a:
                    return new XAttribute(a.Name, a.Value);

                case XElement e:
                    return
                        new XElement(
                            e.Name,
                            e.Attributes().Select(x => Update(x, docPrSequence)),
                            e.Nodes().Select(x => Update(x, docPrSequence)));

                case XText t:
                    return new XText(t.Value);

                default:
                    return xObject;
            }
        }

        /// <summary>
        /// Constructs the part URI from the target URI.
        /// </summary>
        /// <param name="targetUri">The relative target URI.</param>
        /// <returns>
        /// The relative part URI.
        /// </returns>
        [Pure]
        [NotNull]
        static Uri MakePartUri([NotNull] Uri targetUri) => new Uri($"/word/{targetUri}", UriKind.Relative);

        [Pure]
        [NotNull]
        static XAttribute[] Combine([NotNull] IEnumerable<XAttribute> source, [NotNull] IEnumerable<XAttribute> others)
        {
            XAttribute[] attributes = source as XAttribute[] ?? source.ToArray();

            IEnumerable<XAttribute> otherAttributes =
                others.Where(x => !attributes.Any(y => y.Name == x.Name || y.Value == x.Value));

            return attributes.Concat(otherAttributes).ToArray();
        }
    }
}