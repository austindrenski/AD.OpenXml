using System.Linq;
using System.Xml.Linq;
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
                int old =
                    cell.Element(W + "tcPr")?.Element(W + "tcMar")?.Element(W + "start")?.Attribute(W + "w") is XAttribute oldStart
                        ? (int) oldStart
                        : 0;

                cell.Element(W + "tcPr")?.Remove();

                int count =
                    cell.Value.StartsWith("@>")
                        ? cell.Value.Skip(1).TakeWhile(x => x == '>').Count()
                        : 0;

                if (old + count == 0)
                    continue;

                cell.AddFirst(
                    new XElement(W + "tcPr",
                        new XElement(W + "tcMar",
                            new XElement(W + "start",
                                new XAttribute(W + "w", old + count * 144)))));

                foreach (XElement t in cell.Descendants(W + "t"))
                {
                    if (t.Value.StartsWith("@>"))
                        t.Value = new string(t.Value.Skip(1).SkipWhile(x => x == '>').ToArray());
                }
            }

            source.Descendants(W + "trPr").Remove();

            return source;
        }
    }
}