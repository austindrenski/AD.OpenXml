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
            IEnumerable<XElement> headingParagraphFirstRun =
                element.Descendants(W + "p")
                       .Where(x => x.Element(W + "pPr")?
                                    .Element(W + "pStyle")?
                                    .Attribute(W + "val")?
                                    .Value
                                    .Equals("heading1", StringComparison.OrdinalIgnoreCase) ?? false)
                       .Select(x => x.Elements(W + "r").FirstOrDefault())
                       .Where(x => x != null)
                       .ToArray();

            foreach (XElement item in headingParagraphFirstRun)
            {
                item.AddFirst(
                    new XElement(W + "br"));
            }

            return element;
        }
    }
}
