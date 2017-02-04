using System.Xml.Linq;
using JetBrains.Annotations;

namespace AjdExtensions.Html
{
    /// <summary>
    /// Extension methods to transform a &gt;body&lt;...&gt;/body&lt; element into a well-formed HTML document.
    /// </summary>
    [PublicAPI]
    public static class BodyToHtmlExtensions
    {
        /// <summary>
        /// Returns an <see cref="XElement"/> repesenting a well-formed HTML document from the supplied &gt;body&lt;...&gt;/body&lt; element.
        /// </summary>
        /// <param name="element">The &lt;body&gt;...&lt;/body&gt; element.</param>
        /// <param name="title">The name of this HTML document.</param>
        /// <param name="stylesheet">The name, relative path, or absolute path to a CSS stylesheet.</param>
        /// <returns>An <see cref="XElement"/> "html</returns>
        public static XElement BodyToHtml(this XElement element, string title = null, string stylesheet = null)
        {
            XElement html =
                new XElement("html",
                    new XAttribute("lang", "en-US"),
                    new XElement("head",
                        new XElement("meta", 
                            new XAttribute("charset", "UTF-8")
                        ),
                        new XElement("meta",
                            new XAttribute("name", "viewport"),
                            new XAttribute("content", "width=device-width, initial-scale=1.0")
                        ),
                        new XElement("title", title),
                        new XElement("link",
                            new XAttribute("href", stylesheet ?? ""),
                            new XAttribute("type", "text/css"),
                            new XAttribute("rel", "stylesheet")
                        )
                    ),
                    element);
            return html;
        }
    }
}
