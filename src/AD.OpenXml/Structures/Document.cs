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
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        private static readonly XmlWriterSettings XmlWriterSettings =
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
        [NotNull] public static readonly Uri PartName = new Uri("/word/document.xml", UriKind.Relative);

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly string MimeType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

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
                        new XElement(W + "body"));

            XElement documentRelations =
                package.PartExists(DocumentRelsInfo.PartName)
                    ? XElement.Load(package.GetPart(DocumentRelsInfo.PartName).GetStream())
                    : Relationships.Empty;

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
                    .Select(x => new ChartInfo(x.Id, XElement.Load(package.GetPart(new Uri($"/word/{x.Target}", UriKind.Relative)).GetStream())))
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

//            Images =
//                documentRelations
//                    .Elements()
//                    .Select(
//                        x =>
//                            new
//                            {
//                                Id = (string) x.Attribute(DocumentRelsInfo.Attributes.Id),
//                                Target = (string) x.Attribute(DocumentRelsInfo.Attributes.Target)
//                            })
//                    .Where(x => x.Target.StartsWith("media/"))
//                    .Select(x =>
//                    {
//                        MemoryStream ms = new MemoryStream();
//                        Stream s = package.GetPart(new Uri($"/word/{x.Target}", UriKind.Relative)).GetStream();
//                        s.CopyTo(ms);
//                        s.Close();
//                        return ImageInfo.Create(x.Id, x.Target, ms.ToArray());
//                    })
//                    .ToArray();

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
                    .Select(
                        x =>
                            new
                            {
                                x.Id,
                                x.Target,
                                TargetMode =
                                    Enum.TryParse(x.TargetMode, out TargetMode mode)
                                        ? mode
                                        : TargetMode.Internal
                            })
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
            HashSet<XAttribute> attributes =
                new HashSet<XAttribute>(Content.Attributes());

            if (content != null)
            {
                foreach (XAttribute attribute in content.Attributes())
                {
                    if (attributes.Any(x => x.Name == attribute.Name || x.Value == attribute.Value))
                        continue;

                    attributes.Add(attribute);
                }
            }

            XElement document =
                content is null
                    ? Content
                    : new XElement(
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

            foreach (HyperlinkInfo hyperlink in Hyperlinks)
            {
                if (!document.RelationshipExists(hyperlink.RelationId))
                    document.CreateRelationship(hyperlink.Target, hyperlink.TargetMode, HyperlinkInfo.RelationshipType);
            }

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
                        : package.CreatePart(chart.PartName, MimeType);

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
    }
}