using System;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Visits;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
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
        /// <param name="revisionId">
        ///
        /// </param>
        /// <returns></returns>
        public static XElement SetTableStyles(this XElement source, int revisionId)
        {
            if (source.Name != W + "document")
                return source;

            foreach (XElement item in source.Descendants(W + "tblPr"))
            {
                item.RemoveAll();

                if (item.Parent?.Element(W + "tblGrid")?.Elements(W + "gridCol").Count() == 1)
                {
                    item.Add(
                        new XElement(W + "tblStyle",
                            new XAttribute(W + "val", "BoxTable")),
                        new XElement(W + "tblW",
                            new XAttribute(W + "type", "pct"),
                            new XAttribute(W + "w", "5000")),
                        new XElement(W + "tblLook",
                            new XAttribute(W + "val", "04A0"),
                            new XAttribute(W + "firstRow", "1"),
                            new XAttribute(W + "lastRow", "1"),
                            new XAttribute(W + "firstColumn", "0"),
                            new XAttribute(W + "lastColumn", "0"),
                            new XAttribute(W + "noHBand", "1"),
                            new XAttribute(W + "noVBand", "1")));
                }
                else
                {
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
            }

            foreach (XElement cell in source.Descendants(W + "tc"))
            {
                // TODO: this is an ugly way to fix a problem. Too much state being managed. Also...weak heuristic.
                // tc -> tr -> tbl -> tblGrid
                if (cell.Parent?.Parent?.Element(W + "tblGrid")?.Elements(W + "gridCol").Count() == 1)
                {
                    XElement result = DocumentVisit.ExecuteForTableRecursion(cell, revisionId);
                    cell.RemoveNodes();
                    cell.Add(result.Elements());
                }

                if (!cell.Value.Contains("@>"))
                    continue;

                foreach (XElement paragraphWithSymbol in cell.Elements(W + "p").Where(x => x.Value.Contains("@>")))
                {
                    if (paragraphWithSymbol.Element(W + "pPr") is null)
                        paragraphWithSymbol.AddFirst(new XElement(W + "pPr"));

                    foreach (XElement textToIndent in paragraphWithSymbol.Descendants(W + "t").Where(x => x.Value.Contains("@>")))
                    {
                        if (textToIndent.Ancestors(W + "p").First().Element(W + "pPr")?.Element(W + "ind") is null)
                            textToIndent.Ancestors(W + "p").First().Element(W + "pPr")?.Add(new XElement(W + "ind"));

                        XElement indent = textToIndent.Ancestors(W + "p").First().Element(W + "pPr")?.Element(W + "ind");

                        int count = textToIndent.Parent?.Parent?.Value.SkipWhile(x => x != '@').Skip(1).TakeWhile(x => x == '>').Count() ?? 0;

                        if (indent is null)
                            throw new ArgumentException("Indentation symbol code error.");

                        int left = indent.Attribute(W + "left")?.Value.ParseInt() ?? 0;

                        indent.SetAttributeValue(W + "left", left + count * 144);

                        textToIndent.Value = textToIndent.Value.TrimStart('@', '>');
                    }
                }
            }

            //source.Descendants(W + "tcPr").Descendants().Where(x => x.Target != W + "vAlign").Remove();
            source.Descendants(W + "trPr").Remove();
            //source.Descendants(W + "gridCol").Attributes(W + "w").Remove();

            XElement[] tables = source.Element(W + "body")?.Elements(W + "tbl").ToArray() ?? new XElement[0];

            for (int i = 0; i < tables.Length; i++)
            {
                tables[i].Descendants(W + "pPr").Remove();
            }

            return source;
        }
    }
}