using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    [PublicAPI]
    public static class AddLineBreakToHeadingsExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        public static XElement AddLineBreakToHeadings(this XElement element)
        {
            IEnumerable<XElement> headingParagraphProperties =
                element.Descendants(W + "pPr")
                       .Where(x => x.Element(W + "pStyle")?.Attribute(W + "val")?.Value.Equals("heading1", StringComparison.OrdinalIgnoreCase) ?? false)
                       .ToArray();

            foreach (XElement item in headingParagraphProperties)
            {
                item.AddAfterSelf(
                    new XElement(W + "r",
                        new XElement(W + "br")));
            }

            return element;
        }
    }
}
