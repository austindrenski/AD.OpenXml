using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Elements
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class RemoveRsidAttributesExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement RemoveRsidAttributes(this XElement element)
        {
            return element.RemoveAttributesBy(W + "rsidDel")
                          .RemoveAttributesBy(W + "rsidP")
                          .RemoveAttributesBy(W + "rsidR")
                          .RemoveAttributesBy(W + "rsidRDefault")
                          .RemoveAttributesBy(W + "rsidRPr")
                          .RemoveAttributesBy(W + "rsidTr");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="elements"></param>
        /// <returns></returns>
        public static IEnumerable<XElement> RemoveRsidAttributes(this IEnumerable<XElement> elements)
        {
            return elements.Select(x => x.RemoveRsidAttributes());
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="elements"></param>
        /// <returns></returns>
        public static ParallelQuery<XElement> RemoveRsidAttributes(this ParallelQuery<XElement> elements)
        {
            return elements.Select(x => x.RemoveRsidAttributes());
        }
    }
}
