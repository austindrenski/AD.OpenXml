using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.IO.Paths;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Modifies bar chart styling in the target document.
    /// </summary>
    [PublicAPI]
    public static class ModifyPieChartStylesExtensions
    {
        private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        /// Modifies bar chart styling in the target document.
        /// </summary>
        /// <param name="toFilePath"></param>
        public static void ModifyPieChartStyles(this DocxFilePath toFilePath)
        {
            foreach (string item in toFilePath.EnumerateChartPaths())
            {
                XElement element = toFilePath.ReadAsXml(item);
                XElement result = element.ModifyPieChartStyles();
                result.WriteInto(toFilePath, item);
            }
        }

        private static XElement ModifyPieChartStyles(this XElement element)
        {
            if (!element.Descendants(C + "pieChart").Any())
            {
                return element;
            }

            foreach (XElement series in element.Descendants(C + "ser"))
            {
                series.Element(C + "idx")?.SetAttributeValue("val", (string)series.Element(C + "order")?.Attribute("val"));
            }

            element.Descendants(C + "userShapes").Remove();
            element.Descendants(C + "clrMapOvr").Remove();
            element.Descendants().Where(x => x.Name.LocalName == "docPr").Remove();

            element.Element(C + "chart")?.Element(C + "title")?.Remove();

            element =
                element.RemoveByAll(C + "spPr")
                       .RemoveByAll(C + "txPr");

            element.Descendants(C + "legend").Remove();
            element.Descendants(C + "numfmt").Remove();
            element.Descendants(C + "majorGridlines").Remove();
            element.Descendants(A + "endParaRPr").Remove();
            element.Descendants(C + "overlap").Remove();
            element.Descendants(C + "autoTitleDeleted").Remove();
            element.Descendants(C + "noMultiLvlLbl").Remove();

            element.Descendants(C + "varyColors").SetAttributeValues("val", "1");
            element.Descendants(C + "numFmt").SetAttributeValues("sourceLinked", "0");

            element.Add(
                new XElement(C + "spPr",
                    new XElement(A + "ln",
                        new XElement(A + "noFill"))));

            foreach (XElement dataLabels in element.Descendants(C + "dLbls"))
            {
                dataLabels.RemoveAll();
                dataLabels.Add(
                    new XElement(C + "dLblPos",
                        new XAttribute("val", "outEnd")),
                    new XElement(C + "showLegendKey",
                        new XAttribute("val", "1")),
                    new XElement(C + "showVal",
                        new XAttribute("val", "0")),
                    new XElement(C + "showCatName",
                        new XAttribute("val", "1")),
                    new XElement(C + "showSerName",
                        new XAttribute("val", "0")),
                    new XElement(C + "showPercent",
                        new XAttribute("val", "1")),
                    new XElement(C + "showBubbleSize",
                        new XAttribute("val", "0")),
                    new XElement(C + "separator", "\u00A0"),
                    new XElement(C + "showLeaderLines",
                        new XAttribute("val", "1")));
            }

            foreach (XElement paragraphProperties in element.Descendants(A + "pPr"))
            {
                paragraphProperties.RemoveAll();
                paragraphProperties.Add(
                    new XElement(A + "defRPr",
                        new XElement(A + "solidFill",
                            new XElement(A + "prstClr",
                                new XAttribute("val", "black")))));
            }

            return element;
        }
    }
}