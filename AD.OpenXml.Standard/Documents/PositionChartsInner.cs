using System.Linq;
using System.Xml.Linq;
using AD.IO.Standard;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Documents
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class PositionChartsInternalExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="toFilePath"></param>
        public static void PositionChartsInner(this DocxFilePath toFilePath)
        {
            foreach (string item in toFilePath.EnumerateChartPaths())
            {
                XElement element = toFilePath.ReadAsXml(item);

                XElement plotArea = element.Descendants(C + "plotArea").First();

                plotArea.Elements(C + "layout").Remove();

                plotArea.AddFirst(
                    new XElement(C + "layout",
                        new XElement(C + "manualLayout",
                            new XElement(C + "layoutTarget",
                                new XAttribute("val", "inner")),
                            new XElement(C + "xMode",
                                new XAttribute("val", "edge")),
                            new XElement(C + "yMode",
                                new XAttribute("val", "edge")),
                            new XElement(C + "x",
                                new XAttribute("val", "1")),
                            new XElement(C + "y",
                                new XAttribute("val", "1")),
                            new XElement(C + "w",
                                new XAttribute("val", "-0.9")),
                            new XElement(C + "h",
                                new XAttribute("val", "-0.9")))));

                element.WriteInto(toFilePath, item);
            }
        }
    }
}
