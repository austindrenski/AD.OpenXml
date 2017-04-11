using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    [PublicAPI]
    public static class RemoveRsidAttributesExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        public static XElement RemoveRsidAttributes(this XElement element)
        {
            return element.RemoveAttributesBy(W + "rsidP")
                          .RemoveAttributesBy(W + "rsidR")
                          .RemoveAttributesBy(W + "rsidRDefault")
                          .RemoveAttributesBy(W + "rsidRPr")
                          .RemoveAttributesBy(W + "rsidTr");
        }

        public static IEnumerable<XElement> RemoveRsidAttributes(this IEnumerable<XElement> elements)
        {
            return elements.Select(x => x.RemoveRsidAttributes());
        }

        public static ParallelQuery<XElement> RemoveRsidAttributes(this ParallelQuery<XElement> elements)
        {
            return elements.Select(x => x.RemoveRsidAttributes());
        }
    }
}
