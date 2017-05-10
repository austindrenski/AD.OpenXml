using System.Linq;
using System.Xml.Linq;
using AD.Xml.Standard;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Elements
{
    /// <summary>
    /// Extension methods to removes &lt;rPr [...] /&gt; nodes from &lt;pPr [...] /&gt; nodes.
    /// </summary>
    [PublicAPI]
    public static class RemoveRunPropertiesFromParagraphPropertiesExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Removes &lt;rPr [...] /&gt; nodes from &lt;pPr [...] /&gt; nodes.
        /// This method works on the existing <see cref="XElement"/> and returns a reference to it for a fluent syntax.
        /// </summary>
        /// <param name="element">The element to search for descendants.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        public static XElement RemoveRunPropertiesFromParagraphProperties(this XElement element)
        {
            element.Descendants(W + "rPr")
                   .Where(x => x.Parent?.Name == W + "pPr")
                   .Remove();
            return element;
        }
    }
}
