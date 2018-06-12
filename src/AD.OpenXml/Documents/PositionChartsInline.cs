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
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;
        [NotNull] private static readonly XNamespace WP = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;
        // TODO: add to AD.Xml
        [NotNull] private static readonly XNamespace WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

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
                throw new ArgumentNullException(nameof(stream));

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
                throw new ArgumentNullException(nameof(stream));

            MemoryStream result = await stream.CopyPure();

            XElement document = result.ReadXml();

            IEnumerable<XElement> anchors =
                document.Descendants(W + "drawing")
                        .Where(x => x.Elements().FirstOrDefault()?.Name == WP + "anchor")
                        .ToArray();

            foreach (XElement item in anchors)
            {
                item.Element(WP + "anchor")?
                   .AddAfterSelf(
                        new XElement(WP + "inline",
                            new XAttribute("distT", "0"),
                            new XAttribute("distB", "0"),
                            new XAttribute("distL", "0"),
                            new XAttribute("distR", "0"),
                            item.Element(WP + "anchor")?
                                .Elements()
                                .RemoveAttributesBy(WP14 + "anchorId")
                                .RemoveAttributesBy(WP14 + "editId")));

                item.Descendants(WP + "simplePos").Remove();
                item.Descendants(WP + "positionH").Remove();
                item.Descendants(WP + "positionV").Remove();
                item.Descendants(WP + "wrapSquare").Remove();
                item.Descendants(WP + "anchor").Remove();
            }

            return await document.WriteIntoAsync(result, "word/document.xml");
        }
    }
}