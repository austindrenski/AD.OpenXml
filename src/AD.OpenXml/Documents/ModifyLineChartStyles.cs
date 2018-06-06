﻿using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using AD.IO;
using AD.IO.Streams;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Modifies line chart styling in the target document.
    /// </summary>
    [PublicAPI]
    public static class ModifyLineChartStylesExtensions
    {
        private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        /// Modifies line chart styling in the target stream.
        /// </summary>
        /// <param name="stream">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static async Task<MemoryStream> ModifyLineChartStyles([NotNull] [ItemNotNull] this Task<MemoryStream> stream)
        {
            if (stream is null)
                throw new ArgumentNullException(nameof(stream));

            return await ModifyLineChartStyles(await stream);
        }

        /// <summary>
        /// Modifies line chart styling in the target stream.
        /// </summary>
        /// <param name="stream">
        ///
        /// </param>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static async Task<MemoryStream> ModifyLineChartStyles([NotNull] this MemoryStream stream)
        {
            if (stream is null)
                throw new ArgumentNullException(nameof(stream));

            MemoryStream result = await stream.CopyPure();

            foreach (string item in await result.EnumerateChartPartNames())
            {
                result =
                    await result.ReadXml(item)
                                .ModifyLineChartStyles()
                                .WriteIntoAsync(result, item);
            }

            return result;
        }

        /// <summary>
        /// Modifies line chart styling in the target stream.
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        [Pure]
        [NotNull]
        private static XElement ModifyLineChartStyles([NotNull] this XElement element)
        {
            if (element is null)
                throw new ArgumentNullException(nameof(element));

            if (!element.Descendants(C + "lineChart").Any())
                return element;

            foreach (XElement series in element.Descendants(C + "ser"))
            {
                series.Element(C + "idx")?.SetAttributeValue("val", (string) series.Element(C + "order")?.Attribute("val"));
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

            element.Descendants(C + "varyColors").SetAttributeValues("val", "0");
            element.Descendants(C + "gapWidth").SetAttributeValues("val", "150");
            element.Descendants(C + "majorTickMark").SetAttributeValues("val", "out");
            element.Descendants(C + "minorTickMark").SetAttributeValues("val", "none");
            element.Descendants(C + "tickLblPos").SetAttributeValues("val", "low");
            element.Descendants(C + "crosses").SetAttributeValues("val", "autoZero");
            element.Descendants(C + "crossesBetween").SetAttributeValues("val", "between");
            element.Descendants(C + "lblAlgn").SetAttributeValues("val", "ctr");
            element.Descendants(C + "lblOffset").SetAttributeValues("val", "100");

            //element.Add(
            //            new XElement(C + "spPr",
            //                new XElement(A + "ln",
            //                    new XElement(A + "noFill"))));

            //element.Element(C + "chart")?
            //       .Add(
            //            new XElement(C + "legend",
            //                new XElement(C + "legendPos",
            //                    new XAttribute("val", "b")),
            //                new XElement(C + "overlay",
            //                    new XAttribute("val", "0"))));

            //element.Element(C + "chart")?
            //       .Element(C + "plotArea")?
            //       .Add(
            //            new XElement(C + "spPr",
            //                new XElement(A + "noFill"),
            //                new XElement(A + "ln",
            //                    new XElement(A + "solidFill",
            //                        new XElement(A + "prstClr",
            //                            new XAttribute("val", "black"))))));

            //element.Element(C + "chart")?
            //       .Element(C + "plotArea")?
            //       .Element(C + "valAx")?
            //       .AddFirst(
            //            new XElement(C + "spPr",
            //                new XElement(A + "noFill"),
            //                new XElement(A + "ln",
            //                    new XElement(A + "solidFill",
            //                        new XElement(A + "prstClr",
            //                            new XAttribute("val", "black"))))));

            //element.Element(C + "chart")?
            //       .Element(C + "plotArea")?
            //       .Element(C + "catAx")?
            //       .AddFirst(
            //            new XElement(C + "spPr",
            //                new XElement(A + "noFill"),
            //                new XElement(A + "ln",
            //                    new XElement(A + "solidFill",
            //                        new XElement(A + "prstClr",
            //                            new XAttribute("val", "black"))))));

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