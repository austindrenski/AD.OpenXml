using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class TransferFootnotesExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <param name="source"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        public static XElement TransferFootnotes(this XElement element, DocxFilePath source, DocxFilePath result)
        {
            XElement sourceFootnotes;
            try
            {
                sourceFootnotes = source.ReadAsXml("word/footnotes.xml");
            }
            catch
            {
                return element;
            }

            XElement resultFootnotes = result.ReadAsXml("word/footnotes.xml");

            int currentDocumentId =
                resultFootnotes.Descendants(W + "footnote")
                               .Attributes(W + "id")
                               .Select(x => x.Value)
                               .Select(int.Parse)
                               .DefaultIfEmpty(0)
                               .Max();

            IEnumerable<int> fromFootnoteIds =
                sourceFootnotes.Descendants(W + "footnote")
                               .Attributes(W + "id")
                               .Select(x => x.Value)
                               .Where(x => x != "-1" && x != "0")
                               .Select(int.Parse)
                               .ToArray();

            IEnumerable<XElement> footnoteReferenceRunProperties =
                element.Descendants(W + "rPr")
                       .Where(
                           x =>
                               x.Elements(W + "rStyle")
                                .Attributes(W + "val")
                                .Any(y => y.Value.Equals("FootnoteReference", StringComparison.OrdinalIgnoreCase)))
                       .ToArray();

            foreach (XElement runProperty in footnoteReferenceRunProperties)
            {
                runProperty.RemoveAll();
                runProperty.Add(
                    new XElement(
                        W + "rStyle",
                        new XAttribute(
                            W + "val",
                            "FootnoteReference")));
            }

            foreach (int fromId in fromFootnoteIds.OrderByDescending(x => x))
            {
                string toId = $"{currentDocumentId + fromId}";

                element =
                    element.ChangeXAttributeValues(
                        W + "footnoteReference",
                        W + "id",
                        $"{fromId}",
                        toId);
                
                XElement footnote =
                    sourceFootnotes.Elements()
                                   .Single(x => x.Attribute(W + "id")?.Value == $"{fromId}");

                footnote.SetAttributeValue(W + "id", toId);

                footnote.Descendants(W + "p").Attributes().Remove();

                footnote.Descendants(W + "hyperlink").Remove();

                footnote.RemoveRsidAttributes();

                resultFootnotes.Add(footnote);
            }

            resultFootnotes.WriteInto(result, "word/footnotes.xml");

            return element;
        }
    }
}