﻿using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using AD.IO.Streams;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Modifies bar chart styling in the target document.
    /// </summary>
    [PublicAPI]
    public static class ModifyBarChartStylesExtensions
    {
        private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        /// Modifies bar chart styling in the target stream.
        /// </summary>
        /// <param name="stream"></param>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static async Task<MemoryStream> ModifyBarChartStyles([NotNull] this Task<MemoryStream> stream)
        {
            if (stream is null)
                throw new ArgumentNullException(nameof(stream));

            MemoryStream ms = await (await stream).CopyPure();
            Package package = Package.Open(ms);

            foreach (PackagePart part in package.EnumerateChartPartNames())
            {
                using (Stream original = part.GetStream())
                {
                    using (Stream chart = part.GetStream(FileMode.Truncate))
                    {
                        XElement.Load(original)
                                .ModifyBarChartStyles()
                                .Save(chart);
                    }
                }
            }

            package.Close();

            return ms;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        [Pure]
        [NotNull]
        private static XElement ModifyBarChartStyles([NotNull] this XElement element)
        {
            if (element is null)
                throw new ArgumentNullException(nameof(element));

            if (!element.Descendants(C + "barChart").Any())
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
            element.Descendants(C + "gapWidth").SetAttributeValues("val", "150");
            element.Descendants(C + "majorTickMark").SetAttributeValues("val", "out");
            element.Descendants(C + "minorTickMark").SetAttributeValues("val", "none");
            element.Descendants(C + "tickLblPos").SetAttributeValues("val", "low");
            element.Descendants(C + "crosses").SetAttributeValues("val", "autoZero");
            element.Descendants(C + "crossesBetween").SetAttributeValues("val", "between");
            element.Descendants(C + "lblAlgn").SetAttributeValues("val", "ctr");
            element.Descendants(C + "lblOffset").SetAttributeValues("val", "100");
            element.Descendants(C + "numFmt").SetAttributeValues("sourceLinked", "0");

            element.Add(
                new XElement(C + "spPr",
                    new XElement(A + "ln",
                        new XElement(A + "noFill"))));

            element.Element(C + "chart")?
                .Add(
                    new XElement(C + "legend",
                        new XElement(C + "legendPos",
                            new XAttribute("val", "b")),
                        new XElement(C + "overlay",
                            new XAttribute("val", "0"))));

            element.Element(C + "chart")?
                .Element(C + "plotArea")?
                .Add(
                    new XElement(C + "spPr",
                        new XElement(A + "noFill"),
                        new XElement(A + "ln",
                            new XElement(A + "solidFill",
                                new XElement(A + "prstClr",
                                    new XAttribute("val", "black"))))));

            element.Element(C + "chart")?
                .Element(C + "plotArea")?
                .Element(C + "valAx")?
                .AddFirst(
                    new XElement(C + "spPr",
                        new XElement(A + "noFill"),
                        new XElement(A + "ln",
                            new XElement(A + "solidFill",
                                new XElement(A + "prstClr",
                                    new XAttribute("val", "black"))))));

            element.Element(C + "chart")?
                .Element(C + "plotArea")?
                .Element(C + "catAx")?
                .AddFirst(
                    new XElement(C + "spPr",
                        new XElement(A + "noFill"),
                        new XElement(A + "ln",
                            new XElement(A + "solidFill",
                                new XElement(A + "prstClr",
                                    new XAttribute("val", "black"))))));

            foreach (XElement dataLabels in element.Descendants(C + "dLbls"))
            {
                dataLabels.RemoveAll();
                dataLabels.Add(
                    new XElement(C + "showLegendKey",
                        new XAttribute("val", "0")),
                    new XElement(C + "showVal",
                        new XAttribute("val", "0")),
                    new XElement(C + "showCatName",
                        new XAttribute("val", "0")),
                    new XElement(C + "showSerName",
                        new XAttribute("val", "0")),
                    new XElement(C + "showPercent",
                        new XAttribute("val", "0")),
                    new XElement(C + "showBubbleSize",
                        new XAttribute("val", "0")));
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