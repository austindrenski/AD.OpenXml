﻿using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Modifies pie chart styling in the target document.
    /// </summary>
    [PublicAPI]
    public static class ModifyPieChartStylesExtensions
    {
        private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        /// Modifies pie chart styling in the target stream.
        /// </summary>
        /// <param name="package"></param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [NotNull]
        public static Package ModifyPieChartStyles([NotNull] this Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            Package result =
                package.FileOpenAccess.HasFlag(FileAccess.Write)
                    ? package
                    : package.ToPackage(FileAccess.ReadWrite);

            foreach (PackagePart chart in package.EnumerateChartPartNames())
            {
                chart.ReadXml()
                     .ModifyPieChartStyles()
                     .WriteTo(chart);
            }

            return result;
        }

        /// <summary>
        /// Modifies pie chart styling in the target stream.
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        [Pure]
        [NotNull]
        private static XElement ModifyPieChartStyles([NotNull] this XElement element)
        {
            if (element is null)
                throw new ArgumentNullException(nameof(element));

            if (!element.Descendants(C + "pieChart").Any())
                return element;

            foreach (XElement series in element.Descendants(C + "ser"))
            {
                series.Element(C + "idx")?.SetAttributeValue("val", (string) series.Element(C + "order")?.Attribute("val"));
            }

            element =
                element.RemoveByAll(x => x.Name.LocalName == "docPr")
                       .RemoveByAll(x => x.Name == C + "title")
                       .RemoveByAll(C + "spPr")
                       .RemoveByAll(C + "txPr")
                       .RemoveByAll(C + "userShapes")
                       .RemoveByAll(C + "clrMapOvr")
                       .RemoveByAll(C + "legend")
                       .RemoveByAll(C + "numfmt")
                       .RemoveByAll(C + "majorGridlines")
                       .RemoveByAll(A + "endParaRPr")
                       .RemoveByAll(C + "overlap")
                       .RemoveByAll(C + "autoTitleDeleted")
                       .RemoveByAll(C + "noMultiLvlLbl");

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