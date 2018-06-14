using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Structures;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class PositionChartsOuterExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        private static readonly XNamespace WP = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

        /// <summary>
        ///
        /// </summary>
        /// <param name="package"></param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [NotNull]
        public static Package PositionChartsOuter([NotNull] this Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            Package result =
                package.FileOpenAccess.HasFlag(FileAccess.Write)
                    ? package
                    : package.ToPackage(FileAccess.ReadWrite);

            PackagePart part = result.GetPart(Document.PartUri);

            XElement document = part.ReadXml();

            foreach (XElement item in document.Descendants(W + "drawing").Where(x => x.Descendants(C + "chart").Any()))
            {
                item.Element(WP + "inline")?
                   .Element(WP + "extent")?
                   .Remove();

                item.Element(WP + "inline")?
                   .AddFirst(
                        new XElement(WP + "extent",
                            new XAttribute("cx", 914400 * 6.5),
                            new XAttribute("cy", 914400 * 3.5)));
            }

            document.WriteTo(part);

            return result;
        }
    }
}