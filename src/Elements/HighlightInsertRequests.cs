using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    [PublicAPI]
    public static class HighlightInsertRequestsExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        public static XElement HighlightInsertRequests(this XElement element)
        {
            IEnumerable<XElement> inserts = 
                element.Descendants(W + "p")
                       .Where(x => x.Value.Contains('{'));

            foreach (XElement item in inserts.Descendants(W + "r"))
            {
                if (!item.Descendants(W + "rPr").Any())
                {
                    item.AddFirst(new XElement(W + "rPr"));
                }
                item.Element(W + "rPr")?
                    .Add(
                        new XElement(W + "color", 
                            new XAttribute(W + "val", "FF0000")));
            }

            return element;
        }
    }
}
