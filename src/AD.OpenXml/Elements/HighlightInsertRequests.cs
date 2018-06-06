using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class HighlightInsertRequestsExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement HighlightInsertRequests(this XElement element)
        {
            IEnumerable<XElement> appendices =
                element.Descendants(W + "p")
                       .Where(x => x.Value.Contains("{APPENDIX}"));

            foreach (XElement item in appendices.Descendants(W + "r"))
            {
                if (item.Parent?.Element(W + "pPr") is null)
                    item.Parent?.AddFirst(new XElement(W + "pPr"));
                if (item.Parent?.Element(W + "pPr")?.Element(W + "pStyle") is null)
                    item.Parent?.Element(W + "pPr")?.AddFirst(new XElement(W + "pStyle"));

                item.Parent?
                    .Element(W + "pPr")?
                    .Element(W + "pStyle")?.SetAttributeValue(W + "val", "Heading9");

                XElement text = item.Element(W + "t");

                if (text is null)
                    continue;

                text.Value = text.Value.Replace("{", null);
                text.Value = text.Value.Replace("APPENDIX", null);
                text.Value = text.Value.Replace("}", null);
            }

            IEnumerable<XElement> bibliographies =
                element.Descendants(W + "p")
                       .Where(x => x.Value.Contains("{BIBLIOGRAPHY}"));

            foreach (XElement item in bibliographies.Descendants(W + "r"))
            {
                if (item.Parent?.Element(W + "pPr") is null)
                    item.Parent?.AddFirst(new XElement(W + "pPr"));
                if (item.Parent?.Element(W + "pPr")?.Element(W + "pStyle") is null)
                    item.Parent?.Element(W + "pPr")?.AddFirst(new XElement(W + "pStyle"));

                item.Parent?
                    .Element(W + "pPr")?
                    .Element(W + "pStyle")?.SetAttributeValue(W + "val", "PreHeading");

                XElement text = item.Element(W + "t");

                if (text is null)
                    continue;

                text.Value = text.Value.Replace("{", null);
                text.Value = text.Value.Replace("BIBLIOGRAPHY", "Bibliography");
                text.Value = text.Value.Replace("}", null);
            }

            IEnumerable<XElement> inserts =
                element.Descendants(W + "p")
                       .Where(x => x.Value.Contains('{') && !x.Value.Contains("{APPENDIX}") && !x.Value.Contains("{BIBLIOGRAPHY}"));

            foreach (XElement item in inserts.Descendants(W + "r"))
            {
                if (item.Element(W + "rPr") is null)
                    item.AddFirst(new XElement(W + "rPr"));

                item.Element(W + "rPr")?
                    .Add(
                        new XElement(W + "color", 
                            new XAttribute(W + "val", "FF0000")));
            }

            return element;
        }
    }
}
