using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

// ReSharper disable ClassWithVirtualMembersNeverInherited.Global

namespace AD.OpenXml
{
    /// <inheritdoc />
    /// <summary>
    /// Represents an <see cref="OpenXmlVisitor"/> that can transform an OpenXML body node into a well-formed HTML document.
    /// </summary>
    [PublicAPI]
    public class HtmlVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// Represents the 'a:' prefix seen in the markup for chart[#].xml
        /// </summary>
        [NotNull] private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        /// <summary>
        /// Represents the 'c:' prefix seen in the markup for chart[#].xml
        /// </summary>
        [NotNull] private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        // TODO: move into AD.Xml
        /// <summary>
        /// Represents the 'pic:' prefix seen in the markup for 'drawing' elements.
        /// </summary>
        [NotNull] private static readonly XNamespace P = "http://schemas.openxmlformats.org/drawingml/2006/picture";

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of document.xml.
        /// </summary>
        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// The regex to detect heading styles of the case-insensitive form 'heading[0-9]'.
        /// </summary>
        [NotNull] private static readonly Regex HeadingRegex = new Regex("heading(?<level>[0-9])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// The !DOCTYPE declaration.
        /// </summary>
        [NotNull] private static readonly XDocumentType DocumentTypeDeclaration = new XDocumentType("html", null, null, null);

        /// <summary>
        /// The HTML attributes that may be returned.
        /// </summary>
        protected virtual ISet<XName> SupportedAttributes { get; } =
            new HashSet<XName>
            {
                "id",
                "name",
                "class"
            };

        /// <summary>
        /// The HTML elements that may be returned.
        /// </summary>
        protected virtual ISet<XName> SupportedElements { get; } =
            new HashSet<XName>
            {
                "a",
                "b",
                "caption",
                "em",
                "h1",
                "h2",
                "h3",
                "h4",
                "h5",
                "h6",
                "i",
                "p",
                "sub",
                "sup",
                "table",
                "td",
                "th",
                "tr"
            };

        /// <summary>
        /// The mapping between OpenXML names and HTML names.
        /// </summary>
        protected virtual IDictionary<XName, XName> Renames { get; } =
            new Dictionary<XName, XName>
            {
                // @formatter:off
                [W + "body"]     = "article",
                [W + "document"] = "body",
                [W + "drawing"]  = "figure",
                [W + "tbl"]      = "table",
                [W + "tc"]       = "td",
                [W + "val"]      = "class"
                // @formatter:on
            };

        /// <inheritdoc />
        protected override IDictionary<string, XElement> Charts { get; set; }

        /// <inheritdoc />
        protected override IDictionary<string, (string mime, string description, string base64)> Images { get; set; }

        /// <summary>
        /// The 'charset' tage value.
        /// </summary>
        [NotNull]
        protected string CharacterSet { get; } = "utf-8";

        /// <summary>
        /// The 'lang' tage value.
        /// </summary>
        [NotNull]
        protected string Language { get; } = "en";

        /// <summary>
        /// The value for the 'name' attribute on the 'meta' tag.
        /// </summary>
        [NotNull]
        protected string MetaName { get; } = "viewport";

        /// <summary>
        /// The value for the 'content' attribute on the 'meta' tag.
        /// </summary>
        [NotNull]
        protected string MetaContent { get; } = "width=device-width,minimum-scale=1,initial-scale=1";

        /// <inheritdoc />
        protected HtmlVisitor(bool returnOnDefault) : base(returnOnDefault)
        {
        }

        /// <summary>
        /// Returns a new <see cref="HtmlVisitor"/>.
        /// </summary>
        /// <returns>
        /// An <see cref="HtmlVisitor"/>.
        /// </returns>
        [Pure]
        [NotNull]
        public static HtmlVisitor Create()
        {
            return new HtmlVisitor(false);
        }

        ///  <summary>
        ///   Returns an <see cref="XElement"/> repesenting a well-formed HTML document from the supplied w:document node.
        ///  </summary>
        ///  <param name="document">
        ///   The w:document node.
        ///  </param>
        ///  <param name="footnotes">
        ///
        ///  </param>
        ///  <param name="charts">
        ///
        ///  </param>
        /// <param name="images">
        ///
        /// </param>
        /// <param name="title">
        ///   The name of this HTML document.
        ///  </param>
        ///  <param name="stylesheet">
        ///   The name, relative path, or absolute path to a CSS stylesheet.
        ///  </param>
        ///  <returns>
        ///   An <see cref="XElement"/> "html
        ///  </returns>
        ///  <exception cref="ArgumentNullException" />
        [Pure]
        public XObject Visit(XElement document, XElement footnotes, IDictionary<string, XElement> charts, IDictionary<string, (string mime, string description, string base64)> images, string title, string stylesheet = default)
        {
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            Charts = new Dictionary<string, XElement>(charts);

            Images = new Dictionary<string, (string mime, string description, string base64)>(images);

            return
                new XDocument(
                    DocumentTypeDeclaration,
                    new XElement("html",
                        new XAttribute("lang", Language),
                        new XElement("head",
                            new XElement("meta",
                                new XAttribute("charset", CharacterSet)),
                            new XElement("meta",
                                new XAttribute("name", MetaName),
                                new XAttribute("content", MetaContent)),
                            new XElement("title", title),
                            new XElement("link",
                                new XAttribute("href", stylesheet ?? ""),
                                new XAttribute("type", "text/css"),
                                new XAttribute("rel", "stylesheet")),
                            new XElement("style",
                                new XText("article, section { counter-reset: footnote_counter; }"),
                                new XText("footer :target { background: yellow; }"),
                                new XText("a[aria-describedby=\"footnote-label\"]::before { content: '['; }"),
                                new XText("a[aria-describedby=\"footnote-label\"]::after { content: ']'; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { counter-increment: footnote_counter; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { font-size: 0.5em; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { margin-left: 1px; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { text-decoration: none; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { vertical-align: super; }"))),
                        Lift(Visit(document)),
                        Visit(footnotes)));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAnchor(XElement anchor)
        {
            if (anchor is null)
            {
                throw new ArgumentNullException(nameof(anchor));
            }

            return LiftableHelper(anchor);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAreaChart(XElement areaChart)
        {
            if (areaChart is null)
            {
                throw new ArgumentNullException(nameof(areaChart));
            }

            return ChartHelper(areaChart);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAttribute(XAttribute attribute)
        {
            if (attribute is null)
            {
                throw new ArgumentNullException(nameof(attribute));
            }

            XName name = VisitName(attribute.Name);

            return SupportedAttributes.Contains(name) ? new XAttribute(name, attribute.Value) : null;
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitBarChart(XElement barChart)
        {
            if (barChart is null)
            {
                throw new ArgumentNullException(nameof(barChart));
            }

            return ChartHelper(barChart);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitBody(XElement body)
        {
            if (body is null)
            {
                throw new ArgumentNullException(nameof(body));
            }

            return
                new XElement(
                    VisitName(body.Name),
                    Visit(body.Elements()));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitChart(XElement chart)
        {
            if (chart is null)
            {
                throw new ArgumentNullException(nameof(chart));
            }

            XElement chartContent = Charts[(string) chart.Attribute(R + "id")].Element(C + "chart");

            return LiftableHelper(chartContent);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitDocument(XElement document)
        {
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            return
                new XElement(
                    VisitName(document.Name),
                    Visit(document.Elements()));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitDrawing(XElement drawing)
        {
            if (drawing is null)
            {
                throw new ArgumentNullException(nameof(drawing));
            }

            return
                new XElement(
                    VisitName(drawing.Name),
                    // TODO: handle docPr in OpenXmlVisitor dispatch, then override in HtmlVisitor.
//                    new XElement("figcaption", (string) anchor.Element(WP + "docPr")?.Attribute("title")),
//                    new XComment((string) anchor.Element(WP + "docPr")?.Attribute("descr") ?? string.Empty),
                    Visit(drawing.Elements()));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitFootnote(XElement footnote)
        {
            if (footnote is null)
            {
                throw new ArgumentNullException(nameof(footnote));
            }

            string footnoteReference = (string) footnote.Attribute(W + "id");

            return
                new XElement("li",
                    new XAttribute("id", $"footnote_{footnoteReference}"),
                    new XElement("a",
                        new XAttribute("href", $"#footnote_ref_{footnoteReference}"),
                        new XAttribute("aria-label", "Return to content"),
                        Visit(footnote.Nodes())));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitFootnotes(XElement footnotes)
        {
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }

            return
                new XElement("footer",
                    new XAttribute("class", "footnotes"),
                    new XElement("h2",
                        new XAttribute("id", "footnote-label"),
                        new XText("Footnotes")),
                    new XElement("ol",
                        Visit(footnotes.Elements().Where(x => (int) x.Attribute(W + "id") > 0))));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitGraphic(XElement graphic)
        {
            if (graphic is null)
            {
                throw new ArgumentNullException(nameof(graphic));
            }

            return LiftableHelper(graphic);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitGraphicData(XElement graphicData)
        {
            if (graphicData is null)
            {
                throw new ArgumentNullException(nameof(graphicData));
            }

            return LiftableHelper(graphicData);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitInline(XElement inline)
        {
            if (inline is null)
            {
                throw new ArgumentNullException(nameof(inline));
            }

            return LiftableHelper(inline);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitLineChart(XElement lineChart)
        {
            if (lineChart is null)
            {
                throw new ArgumentNullException(nameof(lineChart));
            }

            return ChartHelper(lineChart);
        }

        /// <inheritdoc />
        [Pure]
        protected override XName VisitName(XName name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return Renames.TryGetValue(name, out XName result) ? result : name.LocalName;
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitParagraph(XElement paragraph)
        {
            if (paragraph is null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            XAttribute classAttribute = paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");

            if (classAttribute != null && HeadingRegex.Match((string) classAttribute) is Match match && match.Success)
            {
                return
                    new XElement(
                        $"h{match.Groups["level"].Value}",
                        Visit(paragraph.Attributes()),
                        Visit(paragraph.Value));
            }

            return
                new XElement(
                    VisitName(paragraph.Name),
                    Visit(classAttribute),
                    Visit(paragraph.Attributes()),
                    Visit(paragraph.Nodes()));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitPicture(XElement picture)
        {
            if (picture is null)
            {
                throw new ArgumentNullException(nameof(picture));
            }

            XAttribute imageId = picture.Element(P + "blipFill")?.Element(A + "blip")?.Attribute(R + "embed");

            return
                new XElement("img",
                    new XAttribute("src",
                        Images.TryGetValue((string) imageId, out (string mime, string description, string base64) image)
                            ? $"data:image/{image.mime};base64,{image.base64}"
                            : $"[image: {(string) imageId}]"),
                    new XAttribute("scale", "0"));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitPieChart(XElement pieChart)
        {
            if (pieChart is null)
            {
                throw new ArgumentNullException(nameof(pieChart));
            }

            return ChartHelper(pieChart);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitPlotArea(XElement plotArea)
        {
            if (plotArea is null)
            {
                throw new ArgumentNullException(nameof(plotArea));
            }

            return LiftableHelper(plotArea);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitRun(XElement run)
        {
            if (run is null)
            {
                throw new ArgumentNullException(nameof(run));
            }

            if (run.Element(W + "drawing") is XElement drawing)
            {
                return Visit(drawing);
            }

            if ((string) run.Element(W + "footnoteReference")?.Attribute(W + "id") is string footnoteReference)
            {
                return
                    new XElement("a",
                        new XAttribute("id", $"footnote_ref_{footnoteReference}"),
                        new XAttribute("href", $"#footnote_{footnoteReference}"),
                        new XAttribute("aria-describedby", "footnote-label"),
                        Visit(footnoteReference));
            }

            if ((string) run.Element(W + "rPr")?.Element(W + "vertAlign")?.Attribute(W + "val") == "superscript" ||
                (string) run.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val") == "superscript" ||
                (string) run.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val") == "FootnoteReference")
            {
                return
                    new XElement("sup",
                        Visit(run.Value));
            }

            if ((string) run.Element(W + "rPr")?.Element(W + "vertAlign")?.Attribute(W + "val") == "subscript" ||
                (string) run.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val") == "subscript")
            {
                return
                    new XElement("sub",
                        Visit(run.Value));
            }

            if (run.Element(W + "rPr")?.Element(W + "b") != null ||
                (string) run.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val") == "Strong")
            {
                return
                    new XElement("b",
                        Visit(run.Value));
            }

            if (run.Element(W + "rPr")?.Element(W + "i") != null ||
                (string) run.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val") == "Emphasis")
            {
                return
                    new XElement("i",
                        Visit(run.Value));
            }

            return Visit(run.Value);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitTable(XElement table)
        {
            if (table is null)
            {
                throw new ArgumentNullException(nameof(table));
            }

            XAttribute classAttribute = table.Element(W + "tblPr")?.Element(W + "tblStyle")?.Attribute(W + "val");

            return
                new XElement(
                    VisitName(table.Name),
                    Visit(classAttribute),
                    Visit(table.Attributes()),
                    Visit(table.Nodes()));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitTableCell(XElement cell)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            XAttribute alignment = cell.Elements(W + "p").FirstOrDefault()?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val");

            if (cell.Elements(W + "p").Count() != 1)
            {
                return
                    new XElement(
                        VisitName(cell.Name),
                        Visit(alignment),
                        Visit(cell.Attributes()),
                        Visit(cell.Nodes()));
            }

            XAttribute style = cell.Element(W + "p").Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");

            XAttribute alignmentStyle =
                alignment is null && style is null
                    ? null
                    : alignment is null
                        ? new XAttribute("class", (string) style)
                        : style is null
                            ? new XAttribute("class", (string) alignment)
                            : new XAttribute("class", $"{style} {alignment}");

            return
                new XElement(
                    VisitName(cell.Name),
                    Visit(alignmentStyle),
                    Visit(cell.Attributes()),
                    Visit(cell.Nodes()).Select(LiftSingleton));
        }

        /// <summary>
        /// Constructs a liftable div element.
        /// </summary>
        /// <param name="element">
        /// The <see cref="XElement"/> from which nodes can be lifted.
        /// </param>
        /// <returns>
        /// An <see cref="XObject"/> representing the lift operation.
        /// </returns>
        /// <exception cref="ArgumentNullException" />>
        [Pure]
        [NotNull]
        private XObject LiftableHelper([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            return
                new XElement("div",
                    new XAttribute(Liftable, $"from-{element.Name.LocalName}"),
                    Visit(element.Elements()));
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="chart">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static XObject ChartHelper([NotNull] XElement chart)
        {
            if (chart is null)
            {
                throw new ArgumentNullException(nameof(chart));
            }

            return
                new XElement("data",
                    chart.Elements(C + "ser")
                         .Select(
                             x =>
                                 new XElement("series",
                                     x.Element(C + "tx")?.Element(C + "strRef")?.Element(C + "strCache").Value is string name
                                         ? new XAttribute("name", name)
                                         : null,
                                     (x.Element(C + "cat")?.Element(C + "strRef")?.Element(C + "strCache") ??
                                      x.Element(C + "cat")?.Element(C + "numRef")?.Element(C + "numCache"))?
                                     .Elements(C + "pt")
                                     .Zip(
                                         (x.Element(C + "val")?.Element(C + "strRef")?.Element(C + "strCache") ??
                                          x.Element(C + "val")?.Element(C + "numRef")?.Element(C + "numCache"))?
                                         .Elements(C + "pt"),
                                         (a, b) =>
                                             new XElement("observation",
                                                 new XAttribute("label", a.Value),
                                                 new XAttribute("value", b.Value))))));
        }
    }
}