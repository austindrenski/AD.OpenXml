using System.Linq;
using System.Xml.Linq;
using AD.Xml.Standard;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Elements
{
    /// <summary>
    /// Extension methods to replace &lt;b/&gt; elements with &lt;rStyle val="Strong"/&gt; elements.
    /// </summary>
    [PublicAPI]
    public static class ChangeBoldToStrongExtensions
    {
        private static XNamespace _w = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Replaces &lt;b [val=...]/&gt; descendant elements with &lt;rStyle val="Strong"/&gt; elements.
        /// This method works on the existing <see cref="XElement"/> and returns a reference to it for a fluent syntax.
        /// </summary>
        /// <param name="element">The element to search for descendants.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        public static XElement ChangeBoldToStrong(this XElement element)
        {
            return element.Replace(_w + "b", _w + "rStyle", new XAttribute(_w + "val", "Strong"));
        }
    }
}
