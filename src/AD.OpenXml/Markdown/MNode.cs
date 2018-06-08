using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Markdown
{
    /// <summary>
    /// Represents a Markdown node.
    /// </summary>
    /// <remarks>
    /// See: http://spec.commonmark.org
    /// </remarks>
    [PublicAPI]
    public abstract class MNode
    {
        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull] protected static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public abstract XNode ToHtml();

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public abstract XNode ToOpenXml();
    }
}