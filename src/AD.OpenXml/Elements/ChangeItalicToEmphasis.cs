using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// Extension methods to replace &lt;i/&gt; elements with &lt;rStyle val="Emphasis"/&gt; elements.
    /// </summary>
    [PublicAPI]
    public static class ChangeItalicToEmphasisExtensions
    {
        [NotNull] static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Replaces &lt;i [val=...] /&gt; descendant elements with &lt;rStyle val="Emphasis"/&gt; elements.
        /// This method works on the existing <see cref="XElement"/> and returns a reference to it for a fluent syntax.
        /// </summary>
        /// <param name="element">The element to search for descendants.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        [NotNull]
        public static XElement ChangeItalicToEmphasis([NotNull] this XElement element)
        {
            XElement[] array1 = element.Descendants(W + "i").Where(x => !x.Ancestors(W + "hyperlink").Any()).ToArray();
            XElement[] array2 = array1.Select(x => x.Parent).ToArray();
            array1.Remove();

            foreach (XElement item in array2)
            {
                item.Add(
                    new XElement(W + "rStyle",
                        new XAttribute(W + "val", "Emphasis")));
            }

            return element;
        }
    }
}