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
        /// <param name="stream">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static async Task<MemoryStream> PositionChartsOuter([NotNull] this Task<MemoryStream> stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            return await PositionChartsOuter(await stream);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream">
        ///
        /// </param>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static async Task<MemoryStream> PositionChartsOuter([NotNull] this MemoryStream stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            MemoryStream result = await stream.CopyPure();

            XElement element = result.ReadXml();

            foreach (XElement item in element.Descendants(W + "drawing").Where(x => x.Descendants(C + "chart").Any()))
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

            return await element.WriteIntoAsync(result, "word/document.xml");
        }
    }
}