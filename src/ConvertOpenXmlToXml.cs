using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Extension methods to clean and convert OpenXML nodes into simple XML nodes.
    /// </summary>
    [PublicAPI]
    public static class ConvertOpenXmlToXmlExtensions
    {
        /// <summary>
        /// Transform OpenXML into simplified XML. This includes removing namespaces and most attributes.
        /// This method traverses the XML in a tail-recursive manner. Do not call this method on any element other than the root element.
        /// </summary>
        /// <param name="element">The root element of the XML object being transformed.</param>
        /// <returns>An XElement cleaned of namespaces and attributes.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        public static XElement ConvertOpenXmlToXml(this XElement element)
        {
            return element.RemoveNamespaces()
                          .RemoveRsidAttributes();
        }
    }
}
