using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace AD.OpenXml.Core.Elements
{
    /// <summary>
    /// Set table styles to BlueTableBasic and perform basic cleaning.
    /// </summary>
    [PublicAPI]
    public static class SetTableStylesExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Set table styles to BlueTableBasic and perform basic cleaning.
        /// </summary>
        /// <param name="source">The source content element. This should be the document-node of document.xml.</param>
        /// <returns></returns>
        public static XElement SetTableStyles(this XElement source)
        {
            foreach (XElement item in source.Descendants(W + "tblPr"))
            {
                item.RemoveAll();
                item.Add(
                    new XElement(W + "tblStyle",
                        new XAttribute(W + "val", "BlueTableBasic")),
                    new XElement(W + "tblW",
                        new XAttribute(W + "type", "pct"),
                        new XAttribute(W + "w", "5000")),
                    new XElement(W + "tblLook",
                        new XAttribute(W + "val", "04A0"),
                        new XAttribute(W + "firstRow", "1"),
                        new XAttribute(W + "lastRow", "0"),
                        new XAttribute(W + "firstColumn", "0"),
                        new XAttribute(W + "lastColumn", "0"),
                        new XAttribute(W + "noHBand", "0"),
                        new XAttribute(W + "noVBand", "1")));
            }

            foreach (XElement item in source.Descendants(W + "tc"))
            {
                foreach (XElement text in item.Descendants(W + "t"))
                {
                    text.Value = Regex.Replace(text.Value, "(\\s\\s)+", "");
                    text.Value = Regex.Replace(text.Value, "(\\s\\s)+", "");
                    text.Value = text.Value.Trim();
                }

                if (!item.Value.Contains("@>"))
                {
                    continue;
                }

                foreach (XElement paragraphWithSymbol in item.Elements(W + "p").Where(x => x.Value.Contains("@>")))
                {
                    if (paragraphWithSymbol.Element(W + "pPr") is null)
                    {
                        paragraphWithSymbol.AddFirst(new XElement(W + "pPr"));
                    }

                    foreach (XElement textToIndent in paragraphWithSymbol.Descendants(W + "t").Where(x => x.Value.Contains("@>")))
                    {
                        if (textToIndent.Ancestors(W + "p").First().Element(W + "pPr")?.Element(W + "ind") is null)
                        {
                            textToIndent.Ancestors(W + "p").First().Element(W + "pPr")?.Add(new XElement(W + "ind"));
                        }

                        XElement indent = textToIndent.Ancestors(W + "p").First().Element(W + "pPr")?.Element(W + "ind");

                        int count = textToIndent.Parent?.Parent?.Value.SkipWhile(x => x != '@').Skip(1).TakeWhile(x => x == '>').Count() ?? 0;
                        
                        if (indent is null)
                        {
                            throw new ArgumentException("Indentation symbol code error.");
                        }

                        int left = indent.Attribute(W + "left")?.Value.ParseInt() ?? 0;

                        indent.SetAttributeValue(W + "left", left + count * 144);

                        textToIndent.Value = textToIndent.Value.TrimStart('@', '>');
                    }
                }
            }

            //source.Descendants(W + "tcPr").Descendants().Where(x => x.Name != W + "vAlign").Remove();
            source.Descendants(W + "trPr").Remove();
            //source.Descendants(W + "gridCol").Attributes(W + "w").Remove();

            return source;
        }
    }
}