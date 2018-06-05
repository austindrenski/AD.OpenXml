using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AD.OpenXml.Structures;
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
        /// The regex to detect heading styles of the case-insensitive form 'heading[0-9]'.
        /// </summary>
        [NotNull] private static readonly Regex HeadingRegex = new Regex("heading(?<level>[0-9])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// The regex to detect heading styles of the case-insensitive form 'heading[0-9]'.
        /// </summary>
        [NotNull] private static readonly Regex SequenceRegex = new Regex("SEQ (?<type>[A-z]+) (?<switches>.*) (?<format>\".+\"|[A-z]+ \\\\s [0-9])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// The regex to detect heading styles of the case-insensitive form 'heading[0-9]'.
        /// </summary>
        [NotNull] private static readonly Regex StyleReferenceRegex = new Regex("STYLEREF (?<format>\".+\"|[0-9] \\\\s)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

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
        /// The mapping between OpenXML names and HTML names.
        /// </summary>
        protected virtual IDictionary<XName, XName> Renames { get; } =
            new Dictionary<XName, XName>
            {
                // @formatter:off
                [W + "body"]     = "article",
                [W + "docPr"]    = "figcaption",
                [W + "document"] = "body",
                [W + "drawing"]  = "figure",
                [W + "tbl"]      = "table",
                [W + "tc"]       = "td",
                [W + "val"]      = "class"
                // @formatter:on
            };

        /// <summary>
        /// The mapping of chart id to node.
        /// </summary>
        [NotNull]
        protected virtual IDictionary<string, XElement> Charts { get; }

        /// <summary>
        /// The mapping of image id to data.
        /// </summary>
        [NotNull]
        protected virtual IDictionary<string, (string mime, string description, string base64)> Images { get; }

        /// <summary>
        /// The 'charset' tage value.
        /// </summary>
        [NotNull]
        protected virtual string CharacterSet => "utf-8";

        /// <summary>
        /// The 'lang' tage value.
        /// </summary>
        [NotNull]
        protected virtual string Language => "en";

        /// <summary>
        /// The value for the 'name' attribute on the 'meta' tag.
        /// </summary>
        [NotNull]
        protected virtual string MetaName => "viewport";

        /// <summary>
        /// The value for the 'content' attribute on the 'meta' tag.
        /// </summary>
        [NotNull]
        protected virtual string MetaContent => "width=device-width,minimum-scale=1,initial-scale=1";

        /// <inheritdoc />
        protected HtmlVisitor(bool returnOnDefault) : base(returnOnDefault)
        {
            Charts = new Dictionary<string, XElement>();
            Images = new Dictionary<string, (string mime, string description, string base64)>();
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes an <see cref="HtmlVisitor"/>.
        /// </summary>
        /// <param name="returnOnDefault">True if an element should be returned when handling the default dispatch case.</param>
        /// <param name="charts">Chart data referenced in the content to be visited.</param>
        /// <param name="images">Image data referenced in the content to be visited.</param>
        protected HtmlVisitor(bool returnOnDefault, [NotNull] IDictionary<string, XElement> charts, [NotNull] IDictionary<string, (string mime, string description, string base64)> images) : base(returnOnDefault)
        {
            if (charts is null)
            {
                throw new ArgumentNullException(nameof(charts));
            }

            if (images is null)
            {
                throw new ArgumentNullException(nameof(images));
            }

            Charts = new Dictionary<string, XElement>(charts);
            Images = new Dictionary<string, (string mime, string description, string base64)>(images);
        }

        /// <summary>
        /// Returns an <see cref="XElement"/> repesenting a well-formed HTML document from the supplied w:document node.
        /// </summary>
        /// <param name="document">The w:document node.</param>
        /// <param name="footnotes"></param>
        /// <param name="title">The name of this HTML document.</param>
        /// <param name="stylesheet">The name, relative path, or absolute path to a CSS stylesheet.</param>
        /// <returns>
        /// An <see cref="XElement"/> "html
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public static XObject Visit([NotNull] Document document, [NotNull] Footnotes footnotes, [NotNull] string title, [NotNull] string stylesheet)
        {
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }

            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            if (stylesheet is null)
            {
                throw new ArgumentNullException(nameof(stylesheet));
            }

            return
                new HtmlVisitor(false, document.ChartReferences, document.ImageReferences)
                    .Visit(document.Content, footnotes.Content, title, stylesheet);
        }

        /// <summary>
        /// Returns an <see cref="XElement"/> repesenting a well-formed HTML document from the supplied w:document node.
        /// </summary>
        /// <param name="document">The w:document node.</param>
        /// <param name="footnotes"></param>
        /// <param name="title">The name of this HTML document.</param>
        /// <param name="stylesheet">The name, relative path, or absolute path to a CSS stylesheet.</param>
        /// <returns>
        /// An <see cref="XElement"/> "html
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public XObject Visit([NotNull] XElement document, [NotNull] XElement footnotes, [NotNull] string title, [NotNull] string stylesheet)
        {
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }

            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            if (stylesheet is null)
            {
                throw new ArgumentNullException(nameof(stylesheet));
            }

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
                            new XElement("script",
                                new XAttribute("type", "text/javascript"),
                                new XAttribute("src", "https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.4/latest.js?config=TeX-MML-AM_CHTML"),
                                new XText(string.Empty)),
                            new XElement("style",
                                new XText("article { counter-reset: chapter_counter; }"),
                                new XText("h1 { page-break-before: always; }"),
                                new XText("h1 { counter-increment: chapter_counter; }"),
                                new XText("h1 { counter-reset: footnote_counter figure_counter table_counter; }"),
//                                new XText("h1::before { content:'Chapter ' counter(chapter_counter) '\\A'; white-space: pre; }"),
                                new XText("footer::target { background: yellow; }"),
                                new XText("table { counter-increment: table_counter; }"),
                                new XText("figure { counter-increment: figure_counter; }"),
                                new XText("caption::before { content: 'Table ' counter(chapter_counter) '.' counter(table_counter) ' '; font-weight: bold; }"),
                                new XText("figcaption::before { content: 'Figure ' counter(chapter_counter) '.' counter(figure_counter) ' '; font-weight: bold; }"),
                                new XText("a[aria-describedby=\"footnote-label\"]::after { content: '[' counter(footnote_counter) ']'; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { counter-increment: footnote_counter; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { font-size: 0.5em; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { margin-left: 1px; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { text-decoration: none; }"),
                                new XText("a[aria-describedby=\"footnote-label\"] { vertical-align: super; }")),
                            new XElement("link",
                                new XAttribute("type", "text/css"),
                                new XAttribute("rel", "stylesheet"),
                                new XAttribute("href", stylesheet))),
                        // TODO: dispatch sequences of nodes optionally (e.g. virtual) instead of each node to support section-encapsulation.
                        Lift(Visit(document)),
                        // TODO: handle this as a call at the end of an encapsulated section so that each section can be served as stand alone content.
                        Visit(footnotes)));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAnchor(XElement anchor) => LiftableHelper(anchor);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAreaChart(XElement areaChart) => ChartHelper(areaChart);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAttribute(XAttribute attribute)
            => VisitName(attribute.Name) is XName name &&
               SupportedAttributes.Contains(name)
                   ? new XAttribute(name, attribute.Value)
                   : null;

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitBarChart(XElement barChart) => ChartHelper(barChart);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitBody(XElement body)
            => new XElement(
                VisitName(body.Name),
                Visit(body.Elements()));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitChart(XElement chart)
            => LiftableHelper(Charts[(string) chart.Attribute(R + "id")].Element(C + "chart"));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitDocument(XElement document)
            => new XElement(
                VisitName(document.Name),
                Visit(document.Elements()));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitDocumentProperty(XElement docPr)
            => new XElement(
                VisitName(docPr.Name),
                Visit((string) docPr.Attribute("title") ?? string.Empty));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitDrawing(XElement drawing)
            => new XElement(
                VisitName(drawing.Name),
                Visit(drawing.Elements()));

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
            => new XElement("footer",
                new XAttribute("class", "footnotes"),
                new XElement("h2",
                    new XAttribute("id", "footnote-label"),
                    new XText("Footnotes")),
                // TODO: use <ol> once footnotes are encapsulated by the chapter section.
                new XElement("ul",
                    Visit(footnotes.Elements().Where(x => (int) x.Attribute(W + "id") > 0))));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitGraphic(XElement graphic) => LiftableHelper(graphic);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitGraphicData(XElement graphicData) => LiftableHelper(graphicData);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitInline(XElement inline) => LiftableHelper(inline);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitLineChart(XElement lineChart) => ChartHelper(lineChart);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitMath(XElement math) => MathMLVisitor.Create(true).Visit(math);

        /// <inheritdoc />
        [Pure]
        protected override XName VisitName(XName name)
            => Renames.TryGetValue(name, out XName result)
                   ? result
                   : name.LocalName;

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

            if ((string) classAttribute == "FiguresTablesSourceNote" && !paragraph.Descendants(W + "drawing").Any())
            {
                // Not handled. Greedily subsumed by table nodes.
                return null;
            }

            // ReSharper disable once InvertIf
            if (paragraph.NextNode is XElement next)
            {
                switch ((string) classAttribute)
                {
                    case "CaptionTable" when next.Name == W + "tbl":
                    {
                        // Handled by VisitTable.
                        return null;
                    }
                    case "CaptionFigure" when next.Descendants(W + "drawing").Any():
                    {
                        // Not handled. <figcaption/> created from <w:docPr/>.
                        return null;
                    }
                    default:
                    {
                        break;
                    }
                }
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

            XAttribute imageId = picture.Element(PIC + "blipFill")?.Element(A + "blip")?.Attribute(R + "embed");

            return
                new XElement("img",
                    new XAttribute("scale", "0"),
                    new XAttribute("alt", (string) picture.Element(WP + "docPr")?.Attribute("descr") ?? string.Empty),
                    new XAttribute("src",
                        Images.TryGetValue((string) imageId, out (string mime, string description, string base64) image)
                            ? $"data:image/{image.mime};base64,{image.base64}"
                            : $"[image: {(string) imageId}]"));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitPieChart(XElement pieChart) => ChartHelper(pieChart);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitPlotArea(XElement plotArea) => LiftableHelper(plotArea);

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

            if (run.NextNode is XElement next && next.Element(W + "fldChar") is XElement fieldChar)
            {
                if (fieldChar.Attribute(W + "fldCharType") != null)
                {
                    // Not handled. Field characters are not supported.
                    return null;
                }
            }

            if (run.Element(W + "instrText") != null)
            {
                // Not handled. Field character formatting codes are not supported.
                return null;
            }

            if ((string) run.Element(W + "footnoteReference")?.Attribute(W + "id") is string footnoteReference)
            {
                return
                    new XElement("a",
                        new XAttribute("id", $"footnote_ref_{footnoteReference}"),
                        new XAttribute("href", $"#footnote_{footnoteReference}"),
                        new XAttribute("aria-describedby", "footnote-label"),
                        new XText(string.Empty));
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

            IEnumerable<XObject> caption =
                table.PreviousNode is XElement p && p.Name == W + "p" && (string) p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "CaptionTable"
                    ? p.Nodes()
                    : Enumerable.Empty<XObject>();

            XAttribute classAttribute = table.Element(W + "tblPr")?.Element(W + "tblStyle")?.Attribute(W + "val");

            XObject[] tableNodes = Visit(table.Elements()).ToArray();

            XElement headerRow =
                new XElement("tr",
                    tableNodes.OfType<XElement>()
                              .FirstOrDefault(x => x.Name == "tr")?
                              .Nodes()
                              .Select(
                                  x => !(x is XElement e) || e.Name != "td"
                                           ? x
                                           : new XElement("th",
                                               e.Attributes(),
                                               e.Nodes())));

            IEnumerable<XObject> bodyRows =
                tableNodes.SkipWhile(x => !(x is XElement e) || e.Name != "tr")
                          .Skip(1)
                          .TakeWhile(x => x is XElement e && e.Name == "tr");

            IEnumerable<XElement> footerItems =
                NextWhile(
                    table,
                    x => x as XElement,
                    x => (string) x.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "FiguresTablesSourceNote");

            return
                new XElement(
                    VisitName(table.Name),
                    Visit(classAttribute),
                    Visit(table.Attributes()),
                    new XElement("caption", Visit(caption)),
                    new XElement("thead", headerRow),
                    new XElement("tbody", bodyRows),
                    new XElement("tfoot",
                        new XAttribute("aria-label", "table sources and notes"),
                        footerItems.Select(
                            x =>
                                new XElement("tr",
                                    new XElement("td",
                                        new XAttribute("colspan", headerRow.Elements("th").Count()),
                                        Visit(x.Nodes()))))));
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

            XAttribute style = cell.Element(W + "p")?.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");

            // Lift attributes and content to the cell when the parapgraph is a singleton.
            return
                cell.Elements(W + "p").Count() == 1
                    ? new XElement(
                        VisitName(cell.Name),
                        Visit(MakeClassAttribute(alignment, style)),
                        Visit(cell.Attributes()),
                        Visit(cell.Nodes()).Select(LiftSingleton))
                    : new XElement(
                        VisitName(cell.Name),
                        Visit(alignment),
                        Visit(cell.Attributes()),
                        Visit(cell.Nodes()));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitText(XText text)
            => SequenceRegex.IsMatch(text.Value) ||
               StyleReferenceRegex.IsMatch(text.Value)
                   ? null
                   : text;

        /// <summary>
        ///
        /// </summary>
        /// <param name="chart"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        private static XObject ChartHelper([NotNull] XElement chart)
            => new XElement("data",
                chart.Elements(C + "ser")
                     .Select(
                         x =>
                             new XElement("series",
                                 x.Element(C + "tx")?.Element(C + "strRef")?.Element(C + "strCache")?.Value is string name
                                     ? new XAttribute("name", name)
                                     : null,
                                 (x.Element(C + "cat")?.Element(C + "strRef")?.Element(C + "strCache") ??
                                  x.Element(C + "cat")?.Element(C + "numRef")?.Element(C + "numCache"))?
                                 .Elements(C + "pt")
                                 .Zip(
                                     (x.Element(C + "val")?.Element(C + "strRef")?.Element(C + "strCache") ??
                                      x.Element(C + "val")?.Element(C + "numRef")?.Element(C + "numCache"))?
                                     .Elements(C + "pt")
                                     ??
                                     Enumerable.Empty<XElement>(),
                                     (a, b) =>
                                         new XElement("observation",
                                             new XAttribute("label", a.Value),
                                             new XAttribute("value", b.Value))))));

        /// <summary>
        /// Space delimits the attribute values into a 'class' attribute.
        /// </summary>
        /// <param name="attributes">The <see cref="XAttribute"/> collection to combine.</param>
        /// <returns>
        /// An <see cref="XAttribute"/> with the name 'class' and the values from <paramref name="attributes"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        private static XAttribute MakeClassAttribute([NotNull] [ItemCanBeNull] params XAttribute[] attributes)
            => attributes.Any(x => x != null)
                   ? new XAttribute("class", string.Join(" ", attributes.Where(x => x != null).Select(x => (string) x)))
                   : null;
    }
}