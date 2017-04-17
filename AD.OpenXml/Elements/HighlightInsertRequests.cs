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
            IEnumerable<XElement> appendices =
                element.Descendants(W + "p")
                       .Where(x => x.Value.Contains("{APPENDIX}"));

            foreach (XElement item in appendices.Descendants(W + "r"))
            {
                if (item.Parent?.Element(W + "pPr") is null)
                {
                    item.Parent?.AddFirst(new XElement(W + "pPr"));
                }
                if (item.Parent?.Element(W + "pPr")?.Element(W + "pStyle") is null)
                {
                    item.Parent?.Element(W + "pPr")?.AddFirst(new XElement(W + "pStyle"));
                }

                item.Parent?
                    .Element(W + "pPr")?
                    .Element(W + "pStyle")?.SetAttributeValue(W + "val", "Appendix");

                XElement text = item.Element(W + "t");
                text.Value = text.Value.Replace("{", null);
                text.Value = text.Value.Replace("APPENDIX", null);
                text.Value = text.Value.Replace("}", null);
            }
            
            IEnumerable<XElement> inserts =
                element.Descendants(W + "p")
                       .Where(x => x.Value.Contains('{') && !x.Value.Contains("{APPENDIX}"));

            foreach (XElement item in inserts.Descendants(W + "r"))
            {
                if (item.Element(W + "rPr") is null)
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
