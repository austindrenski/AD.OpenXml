using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO.Standard;
using AD.Xml.Standard;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Documents
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class PositionChartsInlineExtensions
    {
        private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        private static readonly XNamespace D = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

        private static readonly XNamespace Wp2010 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

        /// <summary>
        /// 
        /// </summary>
        /// <param name="toFilePath"></param>
        public static void PositionChartsInline(this DocxFilePath toFilePath)
        {
            IEnumerable<XElement> charts = 
                toFilePath.ReadAsXml()
                          .Descendants(W + "drawing")
                          .Where(x => x.Elements().FirstOrDefault()?.Name == D + "anchor")
                          .ToArray();

            foreach (XElement item in charts)
            {
                item.AddAfterSelf(
                    new XElement(A + "inline",
                        new XAttribute("distT", "0"),
                        new XAttribute("distB", "0"),
                        new XAttribute("distL", "0"),
                        new XAttribute("distR", "0"),
                        item.Element(D + "anchor")?
                            .Elements()
                            .RemoveAttributesBy(Wp2010 + "anchorId")
                            .RemoveAttributesBy(Wp2010 + "editId")));

                item.RemoveBy(D + "anchor");
            }
        }
    }
}
