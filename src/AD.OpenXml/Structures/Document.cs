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
        [NotNull] public static readonly string MimeType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly Uri PartName = new Uri("/word/document.xml", UriKind.Relative);

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
        /// <param name="package">
        /// The package to which changes can be saved.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public Document([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            Content =
                package.PartExists(PartName)
                    ? XElement.Load(package.GetPart(PartName).GetStream())
                    : new XElement(W + "document",
                        new XAttribute(XNamespace.Xmlns + "w", W),
                        new XElement(W + "body"));

            Charts =
                package.GetPart(PartName)
                       .GetRelationshipsByType(ChartInfo.RelationshipType)
                       .Select(
                           x =>
                           {
                               Uri uri = new Uri("/word/" + x.TargetUri, UriKind.Relative);
                               using (Stream stream = package.GetPart(uri).GetStream())
                               {
                                   return new ChartInfo(x.Id, XElement.Load(stream));
                               }
                           })
                       .ToArray();

            Images =
                package.GetPart(PartName)
                       .GetRelationshipsByType(ImageInfo.RelationshipType)
                       .Select(
                           x =>
                           {
                               Uri uri = new Uri("/word/" + x.TargetUri, UriKind.Relative);
                               MemoryStream ms = new MemoryStream();
                               using (Stream s = package.GetPart(uri).GetStream())
                               {
                                   s.CopyTo(ms);
                                   s.Close();
                                   return ImageInfo.Create(x.Id, uri.ToString(), ms.ToArray());
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
        [Pure]
        [NotNull]
        public Document Concat([NotNull] Document other)
        {
            if (other is null)
                throw new ArgumentNullException(nameof(other));

            XElement content = other.Content;
            ChartInfo[] charts = other.Charts as ChartInfo[] ?? other.Charts.ToArray();
            ImageInfo[] images = other.Images as ImageInfo[] ?? other.Images.ToArray();
            HyperlinkInfo[] hyperlinks = other.Hyperlinks as HyperlinkInfo[] ?? other.Hyperlinks.ToArray();

            HashSet<XAttribute> attributes =
                new HashSet<XAttribute>(Content.Attributes());

            foreach (XAttribute attribute in other.Content.Attributes())
            {
                if (attributes.Any(x => x.Name == attribute.Name || x.Value == attribute.Value))
                    continue;

                attributes.Add(attribute);
            }

            XElement document =
                new XElement(
                    Content.Name,
                    attributes,
                    new XElement(
                        W + "body",
                        Content.Element(W + "body")?.Elements(),
                        content.Element(W + "body")?.Elements()));

            document.RemoveDuplicateSectionProperties();

            return
                new Document(
                    document,
                    Charts.Concat(charts),
                    Images.Concat(images),
                    Hyperlinks.Concat(hyperlinks));
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString()
            => $"Document: (Charts: {Charts.Count()}, Images: {Images.Count()}, Hyperlinks: {Hyperlinks.Count()})";

        /// <summary>
        ///
        /// </summary>
        /// <param name="package">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            PackagePart document =
                package.PartExists(PartName)
                    ? package.GetPart(PartName)
                    : package.CreatePart(PartName, MimeType);

            SaveHyperlinks(Hyperlinks, document, Content);

            foreach (ImageInfo image in Images)
            {
                if (!document.RelationshipExists(image.RelationId))
                    document.CreateRelationship(image.Target, TargetMode.Internal, ImageInfo.RelationshipType, image.RelationId);

                PackagePart imagePart =
                    package.PartExists(image.PartName)
                        ? package.GetPart(image.PartName)
                        : package.CreatePart(image.PartName, image.MimeType);

                using (Stream stream = imagePart.GetStream())
                {
                    stream.Write(image.Image.Span);
                }
            }

            foreach (ChartInfo chart in Charts)
            {
                if (!document.RelationshipExists(chart.RelationId))
                    document.CreateRelationship(chart.Target, TargetMode.Internal, ChartInfo.RelationshipType, chart.RelationId);

                PackagePart chartPart =
                    package.PartExists(chart.PartName)
                        ? package.GetPart(chart.PartName)
                        : package.CreatePart(chart.PartName, ChartInfo.MimeType);

                using (Stream stream = chartPart.GetStream())
                {
                    using (XmlWriter xml = XmlWriter.Create(stream, XmlWriterSettings))
                    {
                        chart.Chart.WriteTo(xml);
                    }
                }
            }

            using (Stream stream = document.GetStream())
            {
                using (XmlWriter xml = XmlWriter.Create(stream, XmlWriterSettings))
                {
                    Content.WriteTo(xml);
                }
            }
        }

        private static void SaveHyperlinks(IEnumerable<HyperlinkInfo> hyperlinks, PackagePart document, XElement content)
        {
            (string from, string to)[] updates =
                hyperlinks
                    .Select(
                        x =>
                            new
                            {
                                from = x.RelationId,
                                to = document.RelationshipExists(x.RelationId) &&
                                     document.GetRelationship(x.RelationId).TargetUri == x.Target
                                    ? document.GetRelationship(x.RelationId)
                                    : document.CreateRelationship(x.Target, x.TargetMode, HyperlinkInfo.RelationshipType)
                            })
                    .Select(x => (x.from, x.to.Id))
                    .ToArray();

            XAttribute[] attributes =
                content.DescendantsAndSelf()
                       .Attributes()
                       .Where(x => x.Name == R + "id")
                       .ToArray();

            for (int i = 0; i < attributes.Length; i++)
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
        }
    }
}