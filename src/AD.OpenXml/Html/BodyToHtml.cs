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

        private static readonly HashSet<string> SupportedAttributes =
            new HashSet<string>
            {
                "id",
                "name",
                "val",
                "class"
            };

        private static readonly HashSet<string> SupportedElements =
            new HashSet<string>
            {
                "h1",
                "h2",
                "h3",
                "h4",
                "h5",
                "h6",
                "p",
                "table",
                "td",
                "th",
                "tr",
                "caption"
            };

        private static readonly Dictionary<string, string> Renames =
            new Dictionary<string, string>
            {
                { "tbl", "table" },
                { "tc", "td" }
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
                    new XElement("body", Visit(body)));
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
        private static XElement Visit([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            return
                new XElement(
                    element.Name.LocalName.Rename(),
                    element.VisitAttributes(),
                    element.HasElements ? null : element.Value,
                    element.Elements()
                           .Select(Visit)
                           .Select(VisitHeadings)
                           .Select(VisitParagraphs)
                           .Select(VisitTables));
        }

        /// <summary>
        /// Reconstructs the attributes of the element with only the local name.
        /// </summary>
        /// <param name="element">
        /// The element from which to reconstruct attributes.
        /// </param>
        /// <returns>
        /// The reconstructed attributes.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        private static IEnumerable<XAttribute> VisitAttributes([NotNull] this XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            return
                element.Attributes()
                       .Where(x => !x.IsNamespaceDeclaration)
                       .Where(x => SupportedAttributes.Contains(x.Name.LocalName))
                       .Select(x => new XAttribute(x.Name.LocalName.Rename(), x.Value));
        }

        [Pure]
        [NotNull]
        private static XElement VisitHeadings([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name.LocalName != "p")
            {
                return element;
            }

            string value = (string) element.Element("pPr")?.Element("pStyle")?.Attribute("val");

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
                    element.VisitAttributes(),
                    element.Value);
        }

        [Pure]
        [NotNull]
        private static XElement VisitParagraphs([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name.LocalName != "p")
            {
                return element;
            }

            string alignment = (string) element.Element("pPr")?.Element("jc")?.Attribute("val");

            return
                new XElement(
                    element.Name.LocalName.Rename(),
                    element.VisitAttributes(),
                    alignment != null ? new XAttribute("class", alignment) : null,
                    element.Value);
        }

        [Pure]
        [NotNull]
        private static XElement VisitTables([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name.LocalName != "table")
            {
                return element;
            }

            return
                new XElement(
                    element.Name.LocalName.Rename(),
                    element.VisitAttributes(),
                    element.Elements()
                           .Select(Visit)
                           .Select(VisitTableRows)
                           .Where(x => SupportedElements.Contains(x.Name.LocalName)));
        }

        [Pure]
        [NotNull]
        private static XElement VisitTableRows([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name.LocalName != "tr")
            {
                return element;
            }

            return
                new XElement(
                    element.Name.LocalName.Rename(),
                    element.VisitAttributes(),
                    element.Elements()
                           .Select(Visit)
                           .Select(VisitTableCells));
        }

        [Pure]
        [NotNull]
        private static XElement VisitTableCells([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (element.Name.LocalName != "td")
            {
                return element;
            }

            string alignment = (string) element.Element("p")?.Attribute("class");

            return
                new XElement(
                    element.Name.LocalName.Rename(),
                    element.VisitAttributes(),
                    alignment != null ? new XAttribute("class", alignment) : null,
                    element.Value);
        }

        [Pure]
        [NotNull]
        private static string Rename([NotNull] this string name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return Renames.TryGetValue(name, out string result) ? result : name;
        }
    }
}