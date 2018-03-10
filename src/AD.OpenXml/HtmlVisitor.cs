﻿using System;
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

        /// <summary>
        /// Represents the 'wp:' prefix seen in the markup for 'drawing' elements.
        /// </summary>
        [NotNull] private static readonly XNamespace D = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

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

        /// <inheritdoc />
        /// <summary>
        /// The HTML attributes that may be returned.
        /// </summary>
        protected override ISet<XName> SupportedAttributes { get; } =
            new HashSet<XName>
            {
                "id",
                "name",
                "class"
            };

        /// <inheritdoc />
        /// <summary>
        /// The HTML elements that may be returned.
        /// </summary>
        protected override ISet<XName> SupportedElements { get; } =
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

        /// <inheritdoc />
        /// <summary>
        /// The mapping between OpenXML names and HTML names.
        /// </summary>
        protected override IDictionary<XName, XName> Renames { get; } =
            new Dictionary<XName, XName>
            {
                { "drawing", "figure" },
                { "tbl", "table" },
                { "tc", "td" },
                { "val", "class" }
            };

        /// <inheritdoc />
        protected override IDictionary<string, XElement> Charts { get; set; }

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

        /// <summary>
        /// Returns a new <see cref="HtmlVisitor"/>.
        /// </summary>
        /// <returns>
        /// An <see cref="HtmlVisitor"/>.
        /// </returns>
        public static HtmlVisitor Create()
        {
            return new HtmlVisitor();
        }

        /// <summary>
        ///  Returns an <see cref="XElement"/> repesenting a well-formed HTML document from the supplied w:document node.
        /// </summary>
        /// <param name="document">
        ///  The w:document node.
        /// </param>
        /// <param name="footnotes">
        ///
        /// </param>
        /// <param name="charts">
        ///
        /// </param>
        /// <param name="title">
        ///  The name of this HTML document.
        /// </param>
        /// <param name="stylesheet">
        ///  The name, relative path, or absolute path to a CSS stylesheet.
        /// </param>
        /// <returns>
        ///  An <see cref="XElement"/> "html
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        public XObject Visit(XElement document, XElement footnotes, IDictionary<string, XElement> charts, string title, string stylesheet = default)
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
                        new XElement("body",
                            new XElement("article",
                                Visit(document.Element(W + "body").Nodes()))),
                        new XElement("footer",
                            new XAttribute("class", "footnotes"),
                            new XElement("h2",
                                new XAttribute("id", "footnote-label"),
                                new XText("Footnotes")),
                            new XElement("ol",
                                Visit(footnotes.Elements().Where(x => (int) x.Attribute(W + "id") > 0))))));
        }

        /// <inheritdoc />
        /// <summary>
        ///
        /// </summary>
        /// <param name="element">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override XObject VisitElement(XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            switch (element)
            {
                case XElement e when e.Name.LocalName == "body":
                {
                    return VisitBody(e);
                }
                case XElement e when e.Name.LocalName == "drawing":
                {
                    return VisitDrawing(e);
                }
                case XElement e when e.Name.LocalName == "footnote":
                {
                    return VisitFootnote(e);
                }
                case XElement e when e.Name.LocalName == "p":
                {
                    return VisitParagraph(e);
                }
                case XElement e when e.Name.LocalName == "r":
                {
                    return VisitRun(e);
                }
                case XElement e when e.Name.LocalName == "tbl":
                {
                    return VisitTable(e);
                }
                case XElement e when e.Name.LocalName == "tr":
                {
                    return VisitTableRow(e);
                }
                case XElement e when e.Name.LocalName == "tc":
                {
                    return VisitTableCell(e);
                }
                default:
                {
                    return null;
                }
            }
        }

        /// <inheritdoc />
        /// <summary>
        ///
        /// </summary>
        /// <param name="drawing">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override XObject VisitDrawing(XElement drawing)
        {
            if (drawing is null)
            {
                throw new ArgumentNullException(nameof(drawing));
            }

            XAttribute idAttribute =
                drawing.Element(D + "inline")?.Element(A + "graphic")?.Element(A + "graphicData")?.Element(C + "chart")?.Attribute(R + "id");

            return
                new XElement(
                    VisitName(drawing.Name),
                    Visit(idAttribute),
                    new XElement("div",
                        Charts.TryGetValue((string) idAttribute, out XElement chart)
                            ? (object) chart
                            : new XText($"[figure: {(string) idAttribute}]")),
                    Visit(drawing.Parent?.Elements()?.Where(x => x.Name != W + "drawing")));
        }

        /// <inheritdoc />
        /// <summary>
        ///
        /// </summary>
        /// <param name="footnote">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
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
        /// <summary>
        ///
        /// </summary>
        /// <param name="paragraph">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
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
        /// <summary>
        ///
        /// </summary>
        /// <param name="run">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
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
        /// <summary>
        ///
        /// </summary>
        /// <param name="table">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
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
        /// <summary>
        ///
        /// </summary>
        /// <param name="cell">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
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
    }
}