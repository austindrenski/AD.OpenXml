using System;
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
    ///
    /// </summary>
    [PublicAPI]
    public static class PositionChartsInnerExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        public static async Task<MemoryStream> PositionChartsInner(this Task<MemoryStream> stream)
        {
            if (stream is null)
                throw new ArgumentNullException(nameof(stream));

            return await PositionChartsInner(await stream);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream">
        ///
        /// </param>
        public static async Task<MemoryStream> PositionChartsInner(this MemoryStream stream)
        {
            if (stream is null)
                throw new ArgumentNullException(nameof(stream));

            MemoryStream result = await stream.CopyPure();

            foreach (string item in await result.EnumerateChartPartNames())
            {
                XElement element = result.ReadXml(item);

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

                result = await element.WriteIntoAsync(result, item);
            }

            return result;
        }
    }
}