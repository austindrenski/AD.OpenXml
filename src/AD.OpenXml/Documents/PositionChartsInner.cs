using System;
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
    ///
    /// </summary>
    [PublicAPI]
    public static class PositionChartsInnerExtensions
    {
        [NotNull] private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream"></param>
        public static async Task<MemoryStream> PositionChartsInner(this Task<MemoryStream> stream)
        {
            if (stream is null)
                throw new ArgumentNullException(nameof(stream));

            MemoryStream ms = await (await stream).CopyPure();

            using (Package package = Package.Open(ms))
            {
                foreach (PackagePart part in package.EnumerateChartPartNames())
                {
                    using (Stream chart = part.GetStream())
                    {
                        XElement element = XElement.Load(chart);

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

                        element.Save(chart);
                    }
                }
            }

            return ms;
        }
    }
}