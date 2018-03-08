using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Html
{
    /// <summary>
    /// Extension methods to transform a &gt;body&lt;...&gt;/body&lt; run into a well-formed HTML document.
    /// </summary>
    [PublicAPI]
    public static class BodyToHtmlExtensions
    {
        [NotNull] private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        [NotNull] private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] private static readonly Regex HeadingRegex = new Regex("heading(?<level>[0-9])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        [NotNull] [ItemNotNull] private static readonly HashSet<XName> SupportedAttributes =
            new HashSet<XName>
            {
                "id",
                "name",
                "class"
            };

        [NotNull] [ItemNotNull] private static readonly HashSet<XName> SupportedElements =
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

        [NotNull] private static readonly Dictionary<XName, XName> Renames =
            new Dictionary<XName, XName>
            {
                { "drawing", "figure" },
                { "tbl", "table" },
                { "tc", "td" },
                { "val", "class" }
            };

        /// <summary>
        /// Returns an <see cref="XElement"/> repesenting a well-formed HTML document from the supplied w:body run.
        /// </summary>
        /// <param name="body">The w:body run.</param>
        /// <param name="title">The name of this HTML document.</param>
        /// <param name="stylesheet">The name, relative path, or absolute path to a CSS stylesheet.</param>
        /// <returns>An <see cref="XElement"/> "html</returns>
        public static XElement BodyToHtml(this XElement body, string title = null, string stylesheet = null)
        {
            if (body is null)
            {
                throw new ArgumentNullException(nameof(body));
            }

            return
                new XElement("html",
                    new XAttribute("lang", "en"),
                    new XElement("head",
                        new XElement("meta",
                            new XAttribute("charset", "utf-8")),
                        new XElement("meta",
                            new XAttribute("name", "viewport"),
                            new XAttribute("content", "width=device-width,minimum-scale=1,initial-scale=1")),
                        new XElement("title", title ?? ""),
                        new XElement("link",
                            new XAttribute("href", stylesheet ?? ""),
                            new XAttribute("type", "text/css"),
                            new XAttribute("rel", "stylesheet"))),
                    Visit(body));
        }

        /// <summary>
        /// Visits the node.
        /// </summary>
        /// <param name="xObject">
        /// The XML object to visit.
        /// </param>
        /// <returns>
        /// The visited node.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        private static XObject Visit([CanBeNull] XObject xObject)
        {
            switch (xObject)
            {
                case null:
                {
                    return null;
                }
                case XAttribute a:
                {
                    return VisitAttribute(a);
                }
                case XElement e when e.Name == W + "body":
                {
                    return VisitBody(e);
                }
                case XElement e when e.Name == W + "drawing":
                {
                    return VisitDrawing(e);
                }
                case XElement e when e.Name == W + "p":
                {
                    return VisitParagraph(e);
                }
                case XElement e when e.Name == W + "r":
                {
                    return VisitRun(e);
                }
                case XElement e when e.Name == W + "tbl":
                {
                    return VisitTable(e);
                }
                case XElement e when e.Name == W + "tr":
                {
                    return VisitTableRow(e);
                }
                case XElement e when e.Name == W + "tc":
                {
                    return VisitTableCell(e);
                }
                case XText t:
                {
                    return VisitText(t);
                }
                default:
                {
                    Console.WriteLine($"Skipping unrecognized XObject:\r\n{xObject}");
                    return null;
                }
            }
        }

        [Pure]
        [NotNull]
        private static XName Visit([NotNull] XName name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return Renames.TryGetValue(name.LocalName, out XName result) ? result : name.LocalName;
        }

        [Pure]
        [CanBeNull]
        private static XObject Visit([NotNull] this string text)
        {
            if (text is null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            return Visit(new XText(text));
        }

        /// <summary>
        /// Reconstructs the attribute with only the local name.
        /// </summary>
        /// <param name="attribute">
        /// The attribute to rename.
        /// </param>
        /// <returns>
        /// The reconstructed attribute.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        private static XObject VisitAttribute([NotNull] XAttribute attribute)
        {
            if (attribute is null)
            {
                throw new ArgumentNullException(nameof(attribute));
            }

            XName name = Visit(attribute.Name);

            return SupportedAttributes.Contains(name) ? new XAttribute(name, attribute.Value) : null;
        }

        /// <summary>
        /// Visits the body node.
        /// </summary>
        /// <param name="body">
        /// The body to visit.
        /// </param>
        /// <returns>
        /// The reconstructed attribute.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static XObject VisitBody([NotNull] XElement body)
        {
            if (body is null)
            {
                throw new ArgumentNullException(nameof(body));
            }

            return
                new XElement(
                    Visit(body.Name),
                    body.Nodes().Select(Visit));
        }

        [Pure]
        [NotNull]
        private static XObject VisitDrawing([NotNull] XElement drawing)
        {
            if (drawing is null)
            {
                throw new ArgumentNullException(nameof(drawing));
            }

            XAttribute idAttribute =
                drawing.Element("inline")?
                    .Element(A + "graphic")?
                    .Element(A + "graphicData")?
                    .Element(C + "chart")?
                    .Attribute(R + "id");

            return
                new XElement(
                    Visit(drawing.Name),
                    Visit(idAttribute),
                    Visit($"[figure: {idAttribute?.Value}]"));
        }

        [Pure]
        [CanBeNull]
        private static XObject VisitParagraph([NotNull] XElement paragraph)
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
                        paragraph.Attributes().Select(Visit),
                        Visit(paragraph.Value));
            }

            return
                new XElement(
                    Visit(paragraph.Name),
                    Visit(classAttribute),
                    paragraph.Attributes().Select(Visit),
                    paragraph.Nodes().Select(Visit));
        }

        [Pure]
        [CanBeNull]
        private static XObject VisitRun([NotNull] XElement run)
        {
            if (run is null)
            {
                throw new ArgumentNullException(nameof(run));
            }

            if ((string) run.Element(W + "footnoteReference")?.Attribute(W + "id") is string footnoteReference)
            {
                return
                    new XElement("sup",
                        new XAttribute("class", "footnote_ref"),
                        new XElement("a",
                            new XAttribute("href", $"#footnote_{footnoteReference}"),
                            new XAttribute("id", $"footnote_ref_{footnoteReference}"),
                            Visit(footnoteReference)));
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

        [Pure]
        [NotNull]
        private static XObject VisitTable([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            XAttribute classAttribute = element.Element(W + "tblPr")?.Element(W + "tblStyle")?.Attribute(W + "val");

            return
                new XElement(
                    Visit(element.Name),
                    Visit(classAttribute),
                    element.Attributes().Select(Visit),
                    element.Nodes().Select(Visit));
        }

        [Pure]
        [NotNull]
        private static XElement VisitTableCell([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            XAttribute alignment = element.Elements(W + "p").FirstOrDefault()?.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val");

            if (element.Elements(W + "p").Count() != 1)
            {
                return
                    new XElement(
                        Visit(element.Name),
                        Visit(alignment),
                        element.Attributes().Select(Visit),
                        element.Nodes().Select(Visit));
            }

            XAttribute style = element.Element(W + "p").Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");

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
                    Visit(element.Name),
                    Visit(alignmentStyle),
                    element.Attributes().Select(Visit),
                    element.Nodes().Select(Visit));
        }

        [Pure]
        [NotNull]
        private static XElement VisitTableRow([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            return
                new XElement(
                    Visit(element.Name),
                    element.Attributes().Select(Visit),
                    element.Nodes().Select(Visit));
        }

        [Pure]
        [CanBeNull]
        private static XObject VisitText([NotNull] this XText text)
        {
            if (text is null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            return text;
        }
    }
}