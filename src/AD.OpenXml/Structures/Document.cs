using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
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
        [NotNull] public static readonly string ContentType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly Uri PartName = new Uri("/word/document.xml", UriKind.Relative);

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
                    new XElement(W + "document",
                        new XAttribute(XNamespace.Xmlns + "w", W),
                        new XElement(W + "body"));
            }

            Charts =
                package.GetPart(PartName)
                       .GetRelationshipsByType(ChartInfo.RelationshipType)
                       .Select(
                            x =>
                            {
                                using (Stream s = package.GetPart(PartUri(x.TargetUri)).GetStream())
                                {
                                    return new ChartInfo(x.Id, x.TargetUri, XElement.Load(s));
                                }
                            })
                       .ToArray();

            Images =
                package.GetPart(PartName)
                       .GetRelationshipsByType(ImageInfo.RelationshipType)
                       .Select(
                            x =>
                            {
                                PackagePart part = package.GetPart(PartUri(x.TargetUri));
                                MemoryStream ms = new MemoryStream();
                                using (Stream s = part.GetStream())
                                {
                                    s.CopyTo(ms);
                                    return new ImageInfo(x.Id, x.TargetUri, part.ContentType, ms.ToArray());
                                }
                            })
                       .ToArray();

            Hyperlinks =
                package.GetPart(PartName)
                       .GetRelationshipsByType(HyperlinkInfo.RelationshipType)
                       .Select(x => new HyperlinkInfo(x.Id, x.TargetUri, x.TargetMode))
                       .ToArray();
        }

        ///  <summary>
        ///
        ///  </summary>
        /// <param name="package"></param>
        /// <param name="content"></param>
        /// <param name="charts"></param>
        /// <param name="images"> </param>
        /// <param name="hyperlinks"></param>
        private Document(
            [NotNull] Package package,
            [NotNull] XElement content,
            [NotNull] IEnumerable<ChartInfo> charts,
            [NotNull] IEnumerable<ImageInfo> images,
            [NotNull] IEnumerable<HyperlinkInfo> hyperlinks)
        {
            _package = package;
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
                _package,
                content ?? Content,
                charts ?? Charts,
                images ?? Images,
                hyperlinks ?? Hyperlinks);

        /// <summary>
        ///
        /// </summary>
        /// <param name="other"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public Document Concat([NotNull] Document other)
        {
            if (other is null)
                throw new ArgumentNullException(nameof(other));

            Package result = Package.Open(_package.ToStream(), FileMode.Open);

            PackagePart documentPart = result.GetPart(PartName);

            Dictionary<string, PackageRelationship> resources = new Dictionary<string, PackageRelationship>();

            foreach (ChartInfo info in other.Charts)
            {
                Uri uri =
                    result.PartExists(PartUri(info.TargetUri))
                        ? MakeUniqueUri(info.TargetUri)
                        : info.TargetUri;

                resources[info.Id] = documentPart.CreateRelationship(uri, TargetMode.Internal, ChartInfo.RelationshipType);

                using (Stream stream = result.CreatePart(PartUri(uri), ChartInfo.ContentType).GetStream())
                {
                    using (XmlWriter xml = XmlWriter.Create(stream, XmlWriterSettings))
                    {
                        info.Chart.WriteTo(xml);
                    }
                }
            }

            foreach (ImageInfo info in other.Images)
            {
                Uri uri =
                    result.PartExists(PartUri(info.TargetUri))
                        ? MakeUniqueUri(info.TargetUri)
                        : info.TargetUri;

                resources[info.Id] = documentPart.CreateRelationship(uri, TargetMode.Internal, ImageInfo.RelationshipType);

                using (Stream stream = result.CreatePart(PartUri(uri), info.ContentType).GetStream())
                {
                    stream.Write(info.Image.Span);
                }
            }

            foreach (HyperlinkInfo info in other.Hyperlinks)
            {
                resources[info.Id] = documentPart.CreateRelationship(info.Target, info.TargetMode, HyperlinkInfo.RelationshipType);
            }

            XElement content =
                new XElement(
                    Content.Name,
                    Combine(Content.Attributes(), other.Content.Attributes()),
                    new XElement(
                        W + "body",
                        Content.Element(W + "body")?.Elements(),
                        other.Content.Element(W + "body")?.Elements().Select(x => Update(x, resources))));

            content.RemoveDuplicateSectionProperties();

            using (Stream stream = documentPart.GetStream())
            {
                using (XmlWriter xml = XmlWriter.Create(stream, XmlWriterSettings))
                {
                    content.WriteTo(xml);
                }
            }

            return new Document(result);
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

                Uri uri =
                    target.PartExists(part.Uri)
                        ? MakeUniqueUri(part.Uri)
                        : part.Uri;

                using (Stream from = part.GetStream())
                {
                    using (Stream to = target.CreatePart(uri, part.ContentType).GetStream())
                    {
                        from.CopyTo(to);
                    }
                }
            }
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString()
            => $"(Charts: {Charts.Count()}, Images: {Images.Count()}, Hyperlinks: {Hyperlinks.Count()})";

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
        private static Uri MakeUniqueUri([NotNull] Uri uri)
        {
            ReadOnlySpan<char> original = uri.OriginalString;
            int position = original.LastIndexOf('.');
            ReadOnlySpan<char> left = original.Slice(0, position);
            ReadOnlySpan<char> right = original.Slice(position);

            return new Uri($"{left.ToString()}_{Guid.NewGuid():N}{right.ToString()}", UriKind.Relative);
        }

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
                    when (a.Name == R + "id" || a.Name == R + "embed" || a.Name == "id") &&
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