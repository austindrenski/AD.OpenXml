using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

// ReSharper disable ClassWithVirtualMembersNeverInherited.Global

namespace AD.OpenXml.Html
{
    /// <summary>
    /// Extension methods to transform an OpenXML body node into a well-formed HTML document.
    /// </summary>
    [PublicAPI]
    public class HtmlVisitor
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
        /// The HTML attributes that may be returned.
        /// </summary>
        [NotNull]
        [ItemNotNull]
        protected static ISet<XName> SupportedAttributes { get; } =
            new HashSet<XName>
            {
                "id",
                "name",
                "class"
            };

        /// <summary>
        /// The HTML elements that may be returned.
        /// </summary>
        [NotNull]
        [ItemNotNull]
        protected static ISet<XName> SupportedElements { get; } =
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
        [NotNull]
        protected static IDictionary<XName, XName> Renames { get; } =
            new Dictionary<XName, XName>
            {
                { "drawing", "figure" },
                { "tbl", "table" },
                { "tc", "td" },
                { "val", "class" }
            };

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
        /// Returns an <see cref="XElement"/> repesenting a well-formed HTML document from the supplied w:body run.
        /// </summary>
        /// <param name="body">
        /// The w:body run.
        /// </param>
        /// <param name="title">
        /// The name of this HTML document.
        /// </param>
        /// <param name="stylesheet">
        /// The name, relative path, or absolute path to a CSS stylesheet.
        /// </param>
        /// <returns>
        /// An <see cref="XElement"/> "html
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        public virtual XObject Visit(XElement body, string title, string stylesheet = null)
        {
            if (body is null)
            {
                throw new ArgumentNullException(nameof(body));
            }

            if (title is null)
            {
                throw new ArgumentNullException(nameof(title));
            }

            return
                new XDocument(
                    new XDocumentType("html", null, null, null),
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
                                new XAttribute("rel", "stylesheet"))),
                        Visit(body)));
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
        protected virtual XObject Visit([CanBeNull] XObject xObject)
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

        /// <summary>
        /// Visits the text as an <see cref="XText"/> node.
        /// </summary>
        /// <param name="text">
        /// The text to visit.
        /// </param>
        /// <returns>
        /// A visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject Visit([CanBeNull] string text)
        {
            return Visit(new XText(text));
        }

        /// <summary>
        /// Visits the <see cref="XName"/> for renaming and localization.
        /// </summary>
        /// <param name="name">
        /// The name to visit.
        /// </param>
        /// <returns>
        /// A visited <see cref="XName"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        protected virtual XName Visit([NotNull] XName name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return Renames.TryGetValue(name.LocalName, out XName result) ? result : name.LocalName;
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
        protected virtual XObject VisitAttribute([NotNull] XAttribute attribute)
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
        [CanBeNull]
        protected virtual XObject VisitBody([NotNull] XElement body)
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
        [CanBeNull]
        protected virtual XObject VisitDrawing([NotNull] XElement drawing)
        {
            if (drawing is null)
            {
                throw new ArgumentNullException(nameof(drawing));
            }

            XAttribute idAttribute =
                drawing.Element(D + "inline")?.Element(A + "graphic")?.Element(A + "graphicData")?.Element(C + "chart")?.Attribute(R + "id");

            return
                new XElement(
                    Visit(drawing.Name),
                    Visit(idAttribute),
                    new XElement("div", $"[figure: {idAttribute?.Value}]"),
                    drawing.Parent?.Elements()?.Where(x => x.Name != W + "drawing").Select(Visit));
        }

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
        [CanBeNull]
        protected virtual XObject VisitParagraph([NotNull] XElement paragraph)
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
        [CanBeNull]
        protected virtual XObject VisitRun([NotNull] XElement run)
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
        [CanBeNull]
        protected virtual XObject VisitTable([NotNull] XElement table)
        {
            if (table is null)
            {
                throw new ArgumentNullException(nameof(table));
            }

            XAttribute classAttribute = table.Element(W + "tblPr")?.Element(W + "tblStyle")?.Attribute(W + "val");

            return
                new XElement(
                    Visit(table.Name),
                    Visit(classAttribute),
                    table.Attributes().Select(Visit),
                    table.Nodes().Select(Visit));
        }

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
        [CanBeNull]
        protected virtual XObject VisitTableCell([NotNull] XElement cell)
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
                        Visit(cell.Name),
                        Visit(alignment),
                        cell.Attributes().Select(Visit),
                        cell.Nodes().Select(Visit));
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
                    Visit(cell.Name),
                    Visit(alignmentStyle),
                    cell.Attributes().Select(Visit),
                    cell.Nodes().Select(Visit).Select(LiftSingleton));
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="row">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitTableRow([NotNull] XElement row)
        {
            if (row is null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            return
                new XElement(
                    Visit(row.Name),
                    row.Attributes().Select(Visit),
                    row.Nodes().Select(Visit));
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="text">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitText([NotNull] XText text)
        {
            if (text is null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            return text;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="xObject">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject LiftSingleton([CanBeNull] XObject xObject)
        {
            if (xObject is XContainer container && container.Nodes().Count() <= 1)
            {
                return container.FirstNode;
            }

            return xObject;
        }
    }
}