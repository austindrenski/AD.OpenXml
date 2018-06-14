using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
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
        /// <param name="package"></param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        public static Package PositionChartsInner([NotNull] this Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            Package result =
                package.FileOpenAccess.HasFlag(FileAccess.Write)
                    ? package
                    : package.ToPackage(FileAccess.ReadWrite);

            foreach (PackagePart part in result.EnumerateChartPartNames())
            {
                XElement element = part.ReadXml();

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

                element.WriteTo(part);
            }

            return result;
        }
    }
}