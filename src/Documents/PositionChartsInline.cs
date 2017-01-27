using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    [PublicAPI]
    public static class PositionChartsInlineExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        private static readonly XNamespace D = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

        public static void PositionChartsInline(this DocxFilePath toFilePath)
        {
            IEnumerable<XElement> charts = 
                toFilePath.ReadAsXml()
                          .Descendants(W + "drawing")
                          .Where(x => x.Elements().FirstOrDefault()?.Name == D + "anchor");

            foreach (XElement item in charts)
            {
                item.AddAfterSelf(
                    new XElement(D + "inline",
                        new XAttribute("distT", "0"),
                        new XAttribute("distB", "0"),
                        new XAttribute("distL", "0"),
                        new XAttribute("distR", "0"),
                        item.Element(D + "anchor")?.Elements()));
                item.RemoveBy(D + "anchor");
            }
        }
    }
}
