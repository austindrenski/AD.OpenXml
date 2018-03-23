using System;
using System.Collections.Generic;
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
    // TODO: write a ChartsVisit and document.
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class PositionChartsInlineExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        private static readonly XNamespace WP = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

        private static readonly XNamespace Wp2010 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        public static async Task<MemoryStream> PositionChartsInline(this Task<MemoryStream> stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            return await PositionChartsInline(await stream);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="stream">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        public static async Task<MemoryStream> PositionChartsInline(this MemoryStream stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            MemoryStream result = await stream.CopyPure();

            XElement document = result.ReadXml();

            IEnumerable<XElement> anchors =
                document.Descendants(W + "drawing")
                        .Where(x => x.Elements().FirstOrDefault()?.Name == WP + "anchor")
                        .ToArray();

            foreach (XElement item in anchors)
            {
                item.AddAfterSelf(
                    new XElement(WP + "inline",
                        new XAttribute("distT", "0"),
                        new XAttribute("distB", "0"),
                        new XAttribute("distL", "0"),
                        new XAttribute("distR", "0"),
                        item.Element(WP + "anchor")?
                            .Elements()
                            .RemoveAttributesBy(Wp2010 + "anchorId")
                            .RemoveAttributesBy(Wp2010 + "editId")));

                item.RemoveBy(WP + "anchor");
            }

            return await document.WriteIntoAsync(result, "word/document.xml");
        }
    }
}