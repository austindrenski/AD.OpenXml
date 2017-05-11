using System.Collections.Generic;
using System.Xml.Linq;

namespace AD.OpenXml.Core.Documents
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class PositionChartsOuterExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        private static readonly XNamespace D = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="toFilePath"></param>
        public static void PositionChartsOuter(this DocxFilePath toFilePath)
        {
            XElement element = toFilePath.ReadAsXml();
            IEnumerable<XElement> charts = element.Descendants(W + "drawing");

            foreach (XElement item in charts)
            {
                item.Element(D + "inline")?
                    .Elements(D + "extent")
                    .Remove();

                item.Element(D + "inline")?
                    .AddFirst(
                        new XElement(D + "extent",
                            new XAttribute("cx", 914400 * 6.5),
                            new XAttribute("cy", 914400 * 3.5)));
            }
            element.WriteInto(toFilePath, "word/document.xml");
        }
    }
}
