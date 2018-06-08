using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
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
        #region Fields

        // TODO: move to AD.Xml
        /// <summary>
        /// Represents the prefix 'v:' on VML elements.
        /// </summary>
        [NotNull] private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";

        /// <summary>
        /// The regex to detect heading styles of the case-insensitive form "heading(?[0-9])".
        /// </summary>
        [NotNull] private static readonly Regex HeadingRegex = new Regex("heading(?<level>[0-9])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// The regex to detect sequence references of the case-insensitive form "SEQ (?[A-z]+) (?.*) (?\".+\"|[A-z]+ \\\\s [0-9])".
        /// </summary>
        [NotNull] private static readonly Regex SequenceRegex = new Regex("SEQ (?<type>[A-z]+) (?<switches>.*) (?<format>\".+\"|[A-z]+ \\\\s [0-9])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// The regex to detect style references of the case-insensitive form "STYLEREF (?\".+\"|[0-9] \\\\s)".
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

        #endregion

        #region Constructors

        /// <inheritdoc />
        protected HtmlVisitor(bool allowBaseMethod) : base(allowBaseMethod)
        {
        }

        /// <inheritdoc />
        public HtmlVisitor() : this(false)
        {
            Charts = new Dictionary<string, XElement>();
            Images = new Dictionary<string, (string mime, string description, string base64)>();
        }

        /// <inheritdoc />
        /// <summary>
        /// Initializes an <see cref="HtmlVisitor"/>.
        /// </summary>
        /// <param name="charts">Chart data referenced in the content to be visited.</param>
        /// <param name="images">Image data referenced in the content to be visited.</param>
        public HtmlVisitor(
            [NotNull] IDictionary<string, XElement> charts,
            [NotNull] IDictionary<string, (string mime, string description, string base64)> images)
            : base(false)
        {
            if (charts is null)
                throw new ArgumentNullException(nameof(charts));

            if (images is null)
                throw new ArgumentNullException(nameof(images));

            Charts = new Dictionary<string, XElement>(charts);
            Images = new Dictionary<string, (string mime, string description, string base64)>(images);
        }

        #endregion

        #region Main

        /// <summary>
        /// Returns an <see cref="XObject"/> repesenting a well-formed HTML document from the supplied w:document node.
        /// </summary>
        /// <param name="document">The w:document node.</param>
        /// <param name="footnotes"></param>
        /// <param name="title">The name of this HTML document.</param>
        /// <param name="stylesheetUrl">The name, relative path, or absolute path to a CSS stylesheet.</param>
        /// <param name="styles">The CSS styles to include in a style node.</param>
        /// <returns>
        /// An <see cref="XObject"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public XObject Visit([CanBeNull] XElement document, [CanBeNull] XElement footnotes, [CanBeNull] string title, [CanBeNull] string stylesheetUrl, [CanBeNull] string styles)
            => new XDocument(
                DocumentTypeDeclaration,
                new XElement("html",
                    new XAttribute("lang", Language),
                    new XElement("head",
                        new XElement("meta",
                            new XAttribute("charset", CharacterSet)),
                        new XElement("meta",
                            new XAttribute("name", MetaName),
                            new XAttribute("content", MetaContent)),
                        new XElement("title", title ?? string.Empty),
                        new XElement("script",
                            new XAttribute("type", "text/javascript"),
                            new XAttribute("src", "https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.4/latest.js?config=TeX-MML-AM_CHTML"),
                            new XText(string.Empty)),
                        new XElement("style",
                            @"
h1 {
    counter-reset: footnote_counter;
}


a[aria-describedby='footnote-label'] {
    font-size: 0.5em;
    margin-left: 1px;
    text-decoration: none;
    vertical-align: super;
}

a[aria-describedby='footnote-label']::before {
    content: '[' counter(footnote_counter) ']';
    counter-increment: footnote_counter;
}

a[aria-label='Return to content'] {
    text-decoration: none;
}",
                            styles),
                        new XElement("link",
                            new XAttribute("type", "text/css"),
                            new XAttribute("rel", "stylesheet"),
                            new XAttribute("href", stylesheetUrl ?? string.Empty))),
                    // TODO: dispatch sequences of nodes optionally (e.g. virtual) instead of each node to support section-encapsulation.
                    Lift(Visit(document)),
                    // TODO: handle this as a call at the end of an encapsulated section so that each section can be served as stand alone content.
                    Visit(footnotes)));

        #endregion

        #region Visits

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAnchor(XElement anchor) => MakeLiftable(anchor);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAreaChart(XElement areaChart) => ChartHelper(areaChart);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAttribute(XAttribute attribute)
            => VisitName(attribute.Name) is XName name &&
               SupportedAttributes.Contains(name) &&
               !string.IsNullOrWhiteSpace(attribute.Value)
                   ? new XAttribute(name, attribute.Value)
                   : null;

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitBarChart(XElement barChart) => ChartHelper(barChart);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitChart(XElement chart)
            => MakeLiftable(Charts[(string) chart.Attribute(R + "id")].Element(C + "chart"));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitDocumentProperty(XElement docPr)
            => new XElement(VisitName(docPr.Name), VisitString((string) docPr.Attribute("title")));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitEmbedded(XElement embedded)
        {
            if (!(embedded.Element(V + "shape") is XElement shape))
                return null;

            if (!(shape.Element(V + "imagedata") is XElement imageData))
                return null;

            string altText = (string) shape.Attribute("alt") ?? string.Empty;
            string rId = (string) imageData.Attribute(R + "id");

            return
                new XElement("img",
                    new XAttribute("scale", "0"),
                    new XAttribute("alt", altText),
                    new XAttribute("src", string.Empty),
                    new XElement("span",
                        new XAttribute("class", "error"),
                        Images.TryGetValue(rId, out (string mime, string description, string base64) image)
                            ? $"Images in '{image.mime}' format are not supported: {rId}."
                            : $"An embedded image was detected but not found: {rId}."));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitFootnote(XElement footnote)
        {
            int footnoteReference = (int) footnote.Attribute(W + "id");

            return
                footnoteReference > 0
                    ? new XElement("li",
                        new XAttribute("id", $"footnote_{footnoteReference}"),
                        new XElement("a",
                            new XAttribute("href", $"#footnote_ref_{footnoteReference}"),
                            new XAttribute("aria-label", "Return to content"),
                            Visit(footnote.Nodes())))
                    : null;
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitFootnotes(XElement footnotes)
            => new XElement("footer",
                new XAttribute("class", "footnotes"),
                new XElement("h2",
                    new XAttribute("id", "footnote-label"),
                    new XText("Footnotes")),
                new XElement("ol",
                    Visit(footnotes.Nodes())));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitGraphic(XElement graphic) => MakeLiftable(graphic);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitGraphicData(XElement graphicData) => MakeLiftable(graphicData);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitInline(XElement inline) => MakeLiftable(inline);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitLineChart(XElement lineChart) => ChartHelper(lineChart);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitMath(XElement math) => MathMLVisitor.Create(true).Visit(math);

        /// <inheritdoc />
        [Pure]
        protected override XName VisitName(XName name) => Renames.TryGetValue(name, out XName result) ? result : name.LocalName;

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitParagraph(XElement paragraph)
        {
            XAttribute classAttribute = ParagraphStyle(paragraph);

            if (classAttribute is null)
                return base.VisitParagraph(paragraph);

            // TODO: can we handle list paragraphs?
//            if ((string) classAttribute == "ListParagraph")
//            {
//                // currently building a list
//                if (PreviousParagraphStyleEquals(paragraph, "ListParagraph"))
//                {
//                    return
//                        new XElement("li",
//                            Visit(paragraph.Attributes()),
//                            Visit(paragraph.Nodes()));
//                }
//
//                // start building a list
//                return
//                    new XElement("ol",
//                        Visit(paragraph.Attributes()),
//                        new XElement("li",
//                            Visit(paragraph.Attributes()),
//                            Visit(paragraph.Nodes())),
//                        Visit(NextWhile(paragraph, x => x is XElement e && "ListParagraph" == (string) ParagraphStyle(e))));
//            }

            if (HeadingRegex.Match((string) classAttribute) is Match match && match.Success)
            {
                return
                    new XElement($"h{match.Groups["level"].Value}",
                        Visit(paragraph.Attributes()),
                        VisitString((string) paragraph));
            }

            // Not handled. Greedily subsumed by table nodes.
            if ((string) classAttribute == "FiguresTablesSourceNote" && !paragraph.Descendants(W + "drawing").Any())
                return null;

            // ReSharper disable once InvertIf
            if (paragraph.NextNode is XElement next)
            {
                switch ((string) classAttribute)
                {
                    // Handled by VisitTable.
                    case "CaptionTable" when next.Name == W + "tbl":
                        return null;

                    // Not handled. <figcaption/> created from <w:docPr/>.
                    case "CaptionFigure" when next.Descendants(W + "drawing").Any():
                        return null;

                    default:
                        break;
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
        protected override XObject VisitPlotArea(XElement plotArea) => MakeLiftable(plotArea);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitRun(XElement run)
        {
            if (run.Element(W + "drawing") is XElement drawing)
                return Visit(drawing);

            if (run.Element(W + "object") is XElement embedded)
                return Visit(embedded);

            if (run.NextNode is XElement next && next.Element(W + "fldChar") is XElement fieldChar)
            {
                // Not handled. Field characters are not supported.
                if (fieldChar.Attribute(W + "fldCharType") != null)
                    return null;
            }

            // Not handled. Field character formatting codes are not supported.
            if (run.Element(W + "instrText") != null)
                return null;

            if ((string) run.Element(W + "footnoteReference")?.Attribute(W + "id") is string footnoteReference)
            {
                return
                    new XElement("a",
                        new XAttribute("id", $"footnote_ref_{footnoteReference}"),
                        new XAttribute("href", $"#footnote_{footnoteReference}"),
                        new XAttribute("aria-describedby", "footnote-label"),
                        new XText(string.Empty));
            }

            XElement rPr = run.Element(W + "rPr");

            if ("superscript" == (string) rPr?.Element(W + "vertAlign")?.Attribute(W + "val") ||
                "superscript" == (string) rPr?.Element(W + "rStyle")?.Attribute(W + "val") ||
                "FootnoteReference" == (string) rPr?.Element(W + "rStyle")?.Attribute(W + "val"))
                return new XElement("sup", VisitString((string) run));

            if ("subscript" == (string) rPr?.Element(W + "vertAlign")?.Attribute(W + "val") ||
                "subscript" == (string) rPr?.Element(W + "rStyle")?.Attribute(W + "val"))
                return new XElement("sub", VisitString((string) run));

            if (null != rPr?.Element(W + "b") || "Strong" == (string) rPr?.Element(W + "rStyle")?.Attribute(W + "val"))
                return new XElement("b", VisitString((string) run));

            if (null != rPr?.Element(W + "i") || "Emphasis" == (string) rPr?.Element(W + "rStyle")?.Attribute(W + "val"))
                return new XElement("i", VisitString((string) run));

            return VisitString((string) run);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitTable(XElement table)
        {
            IEnumerable<XNode> caption =
                table.PreviousNode is XElement p &&
                W + "p" == p.Name &&
                "CaptionTable" == (string) p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")
                    ? p.Nodes()
                    : Enumerable.Empty<XNode>();

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
                                            : new XElement("th", e.Attributes(), e.Nodes())));

            IEnumerable<XObject> bodyRows =
                tableNodes.SkipWhile(x => !(x is XElement e) || e.Name != "tr")
                          .Skip(1)
                          .TakeWhile(x => x is XElement e && e.Name == "tr");

            IEnumerable<XElement> footerItems =
                NextWhile(
                        table,
                        x => x is XElement e && (string) e.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "FiguresTablesSourceNote")
                   .OfType<XElement>();

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
            XAttribute alignment = cell.Elements(W + "p").FirstOrDefault()?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val");
            XAttribute style = cell.Element(W + "p")?.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");

            // Lift attributes and content to the cell when the parapgraph is a singleton.
            return
                cell.Elements(W + "p").Count() == 1
                    ? new XElement(
                        VisitName(cell.Name),
                        Visit(new XAttribute("class", string.Join(" ", (string) alignment, (string) style).Trim())),
                        Visit(cell.Attributes()),
                        Visit(cell.Nodes()).Select(x => x is XContainer c ? LiftSingleton(c) : x))
                    : new XElement(
                        VisitName(cell.Name),
                        Visit(new XAttribute("class", (string) alignment ?? string.Empty)),
                        Visit(cell.Attributes()),
                        Visit(cell.Nodes()));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitText(XText text)
            => SequenceRegex.IsMatch(text.Value) || StyleReferenceRegex.IsMatch(text.Value) ? null : text;

        #endregion

        #region Helpers

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
        ///
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        private static XAttribute ParagraphStyle([CanBeNull] XElement paragraph)
            => paragraph?.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");

        /// <summary>
        ///
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="style"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        private static bool NextParagraphStyleEquals([CanBeNull] XElement paragraph, [CanBeNull] string style)
            => paragraph?.NextNode is XElement e && W + "p" == e.Name && style == (string) ParagraphStyle(e);

        /// <summary>
        ///
        /// </summary>
        /// <param name="paragraph"></param>
        /// <param name="style"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        private static bool PreviousParagraphStyleEquals([CanBeNull] XElement paragraph, [CanBeNull] string style)
            => paragraph?.PreviousNode is XElement e && W + "p" == e.Name && style == (string) ParagraphStyle(e);

        #endregion
    }
}