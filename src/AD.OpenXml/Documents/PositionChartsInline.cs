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
        private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        private static readonly XNamespace D = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

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

            XElement document = result.ReadAsXml();

            IEnumerable<XElement> charts =
                document.Descendants(W + "drawing")
                        .Where(x => x.Elements().FirstOrDefault()?.Name == D + "anchor")
                        .ToArray();

            foreach (XElement item in charts)
            {
                item.AddAfterSelf(
                    new XElement(A + "inline",
                        new XAttribute("distT", "0"),
                        new XAttribute("distB", "0"),
                        new XAttribute("distL", "0"),
                        new XAttribute("distR", "0"),
                        item.Element(D + "anchor")?
                            .Elements()
                            .RemoveAttributesBy(Wp2010 + "anchorId")
                            .RemoveAttributesBy(Wp2010 + "editId")));

                item.RemoveBy(D + "anchor");
            }

            return await document.WriteInto(result, "word/document.xml");
        }
    }
}