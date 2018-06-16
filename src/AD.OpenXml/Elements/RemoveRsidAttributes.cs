using System;
using System.Collections.Generic;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// Provides extension methods to remove Revision Save IDs (rsid) attributes from XML trees.
    /// </summary>
    [PublicAPI]
    public static class RemoveRsidAttributesExtensions
    {
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;
        [NotNull] private static readonly XName Rsid = W + "rsid";
        [NotNull] private static readonly XName RsidDel = W + "rsidDel";
        [NotNull] private static readonly XName RsidP = W + "rsidP";
        [NotNull] private static readonly XName RsidR = W + "rsidR";
        [NotNull] private static readonly XName RsidRDefault = W + "rsidRDefault";
        [NotNull] private static readonly XName RsidRPr = W + "rsidRPr";
        [NotNull] private static readonly XName RsidSect = W + "rsidSect";
        [NotNull] private static readonly XName RsidTr = W + "rsidTr";

        /// <summary>
        ///
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public static XElement RemoveRsidAttributes([NotNull] this XElement element)
        {
            if (element is null)
                throw new ArgumentNullException(nameof(element));

            return
                element.RemoveAttributesBy(Rsid)
                       .RemoveAttributesBy(RsidDel)
                       .RemoveAttributesBy(RsidP)
                       .RemoveAttributesBy(RsidR)
                       .RemoveAttributesBy(RsidRDefault)
                       .RemoveAttributesBy(RsidRPr)
                       .RemoveAttributesBy(RsidSect)
                       .RemoveAttributesBy(RsidTr);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="elements"></param>
        /// <returns></returns>
        [NotNull]
        [LinqTunnel]
        [ItemNotNull]
        [CollectionAccess(CollectionAccessType.Read)]
        public static IEnumerable<XElement> RemoveRsidAttributes([NotNull] [ItemCanBeNull] this IEnumerable<XElement> elements)
        {
            if (elements is null)
                throw new ArgumentNullException(nameof(elements));

            foreach (XElement element in elements)
            {
                if (element is null)
                    continue;

                if (element.RemoveRsidAttributes() is XElement result)
                    yield return result;
            }
        }
    }
}