using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// Extension methods to replace &lt;vertAlign val="superscript" /&gt; elements with &lt;rStyle val="FootnoteReference" /&gt; elements.
    /// </summary>
    [PublicAPI]
    public static class ChangeSuperscriptToReferenceExtensions
    {
        [NotNull] static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Replaces &lt;vertAlign val="superscript" /&gt; descendant elements with &lt;rStyle val="FootnoteReference" /&gt; elements.
        /// This method works on the existing <see cref="XElement"/> and returns a reference to it for a fluent syntax.
        /// </summary>
        /// <param name="element">The element to search for descendants.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        [NotNull]
        public static XElement ChangeSuperscriptToReference([NotNull] this XElement element)
        {
            IEnumerable<XElement> superscriptItems =
                element.Descendants(W + "vertAlign")
                       .Where(x => x.Attribute(W + "val")?.Value == "superscript")
                       .ToArray();

            foreach (XElement item in superscriptItems)
            {
                item.AddAfterSelf(new XElement(W + "rStyle", new XAttribute(W + "val", "FootnoteReference")));
                item.Remove();
            }

            return element;
        }
    }
}