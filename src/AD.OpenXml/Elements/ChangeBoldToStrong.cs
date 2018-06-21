using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// Extension methods to replace &lt;b/&gt; elements with &lt;rStyle val="Strong"/&gt; elements.
    /// </summary>
    [PublicAPI]
    public static class ChangeBoldToStrongExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Replaces &lt;b [val=...]/&gt; descendant elements with &lt;rStyle val="Strong"/&gt; elements.
        /// This method works on the existing <see cref="XElement"/> and returns a reference to it for a fluent syntax.
        /// </summary>
        /// <param name="element">The element to search for descendants.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        public static XElement ChangeBoldToStrong(this XElement element)
            => element.Replace(
                W + "b",
                W + "rStyle",
                new XAttribute(W + "val", "Strong"));
    }
}
