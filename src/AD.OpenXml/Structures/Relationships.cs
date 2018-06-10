using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    // TODO: document ContentTypes
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public class Relationships
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly XElement Empty = new XElement(P + "Relationships");

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly string MimeType = "application/vnd.openxmlformats-package.relationships+xml";
    }
}