using System;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Provides methods to open a Microsoft Word document as an XElement.
    /// </summary>
    [PublicAPI]
    public static class AsXmlExtensions
    {
        /// <summary>
        /// Transform OpenXML into simplified XML. This includes removing namespaces and most attributes.
        /// This method traverses the XML in a tail-recursive manner.
        /// </summary>
        /// <param name="element">
        /// The root element of the XML object being transformed.
        /// </param>
        /// <returns>
        /// An XElement cleaned of namespaces and attributes.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public static XElement AsXml([NotNull] this XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            return
                new XElement(
                    element.Name.LocalName,
                    element.Attributes().Select(AsXml),
                    element.Elements().Select(AsXml),
                    element.HasElements ? null : (string) element);
        }

        [Pure]
        [CanBeNull]
        private static XObject AsXml([NotNull] this XAttribute attribute)
        {
            if (attribute is null)
            {
                throw new ArgumentNullException(nameof(attribute));
            }

            switch (attribute.Name.LocalName)
            {
                case "fileName":
                case "fldCharType:":
                {
                    return new XAttribute(attribute.Name.LocalName, attribute.Value);
                }
                case "val":
                {
                    return new XText((string) attribute);
                }
                default:
                {
                    return null;
                }
            }
        }
    }
}