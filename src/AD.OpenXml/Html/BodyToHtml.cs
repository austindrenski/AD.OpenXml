using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Html
{
    /// <summary>
    /// Extension methods to transform a &gt;body&lt;...&gt;/body&lt; element into a well-formed HTML document.
    /// </summary>
    [PublicAPI]
    public static class BodyToHtmlExtensions
    {
        private static readonly Regex HeadingRegex = new Regex("heading(?<level>[0-9])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly HashSet<XName> SupportedAttributes =
            new HashSet<XName>
            {
                "id",
                "name",
                "class"
            };

        private static readonly HashSet<XName> SupportedElements =
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

        private static readonly Dictionary<XName, XName> Renames =
            new Dictionary<XName, XName>
            {
                { "tbl", "table" },
                { "tc", "td" },
                { "val", "class" }
            };

        /// <summary>
        /// Returns an <see cref="XElement"/> repesenting a well-formed HTML document from the supplied w:body element.
        /// </summary>
        /// <param name="body">The w:body element.</param>
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
        /// <param name="node">
        /// The node to visit.
        /// </param>
        /// <returns>
        /// The visited node.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static XObject Visit([NotNull] XNode node)
        {
            if (node is null)
            {
                throw new ArgumentNullException(nameof(node));
            }

            switch (node)
            {
                case XText text:
                {
                    return Visit(text);
                }
                case XElement e when e.Name == "p:" :
                {
                    return Visit(element);
                }
                default:
                {
                    throw new NotImplementedException();
                }
            }
        }

        /// <summary>
        /// Reconstructs the element with only the local name.
        /// </summary>
        /// <param name="element">
        /// The element to reconstruct.
        /// </param>
        /// <returns>
        /// The reconstructed element.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static XObject Visit([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            return
                new XElement(
                    element.Name.Visit(),
                    element.Attributes().Select(Visit),
                    element.HasElements ? null : element.Value,
                    element.Elements()
                           .Select(Visit)
                           .Select(VisitHeadings)
                           .Select(VisitParagraphs)
                           .Select(VisitTables)
                           .Select(VisitTableRows)
                           .Select(VisitTableCells));
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
        private static XObject Visit([NotNull] this XAttribute attribute)
        {
            if (attribute is null)
            {
                throw new ArgumentNullException(nameof(attribute));
            }

            XName name = attribute.Name.Visit();

            return SupportedAttributes.Contains(name) ? new XAttribute(name, attribute.Value) : null;
        }

        [Pure]
        [NotNull]
        private static XName Visit([NotNull] this XName name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return Renames.TryGetValue(name.LocalName, out XName result) ? result : name.LocalName;
        }

        [Pure]
        [NotNull]
        private static XObject Visit([NotNull] this XText text)
        {
            if (text is null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            return text;
        }

        [Pure]
        [NotNull]
        private static XObject VisitHeadings([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name != "p")
            {
                return element;
            }

            string value = (string) element.Element("pPr")?.Element("pStyle")?.Attribute("class");

            if (value is null)
            {
                return element;
            }

            Match match = HeadingRegex.Match(value);

            if (!match.Success)
            {
                return element;
            }

            return
                new XElement(
                    $"h{match.Groups["level"].Value}",
                    element.Attributes(),
                    new XText(element.Value));
        }

        [Pure]
        [NotNull]
        private static XObject VisitParagraphs([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name != "p")
            {
                return element;
            }

            return
                new XElement(
                    element.Name,
                    element.Attributes(),
                    element.Element("pPr")?.Element("pStyle")?.Attribute("class"),
                    element.Elements()
                           .Select(VisitRuns)
                           .Where(x => x is XText || x is XElement y && SupportedElements.Contains(y.Name)));
        }

        [Pure]
        [NotNull]
        private static XObject VisitRuns([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name != "r")
            {
                return element;
            }

            if ((string) element.Element("footnoteReference")?.Attribute("id") is string footnoteReference)
            {
                return
                    new XElement("sup",
                        new XElement("a",
                            new XAttribute("href", $"#footnote{footnoteReference}"),
                            new XAttribute("id", $"r{footnoteReference}"),
                            $"[{footnoteReference}]"));
            }


            if (element.Element("rPr")?.Element("b") != null || (string) element.Element("rPr")?.Element("rStyle")?.Attribute("class") == "Strong")
            {
                return
                    new XElement("b",
                        new XText(element.Value));
            }

            if (element.Element("rPr")?.Element("i") != null || (string) element.Element("rPr")?.Element("rStyle")?.Attribute("class") == "Emphasis")
            {
                return
                    new XElement("i",
                        new XText(element.Value));
            }

            return new XText(element.Value);
        }

        [Pure]
        [NotNull]
        private static XObject VisitTables([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name != "table")
            {
                return element;
            }

            return
                new XElement(
                    element.Name,
                    element.Attributes(),
                    element.Elements().Where(x => SupportedElements.Contains(x.Name)));
        }

        [Pure]
        [NotNull]
        private static XObject VisitTableRows([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name != "tr")
            {
                return element;
            }

            return
                new XElement(
                    element.Name,
                    element.Attributes(),
                    element.Elements().Where(x => SupportedElements.Contains(x.Name)));
        }

        [Pure]
        [NotNull]
        private static XObject VisitTableCells([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name != "td")
            {
                return element;
            }

            return
                new XElement(
                    element.Name,
                    element.Attributes(),
                    element.Element("p")?.Attribute("class"),
                    new XText(element.Value));
        }
    }
}