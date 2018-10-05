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
        [NotNull] static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Set table styles to BlueTableBasic and perform basic cleaning.
        /// </summary>
        /// <param name="source">The source content element. This should be the document-node of document.xml.</param>
        /// <param name="revisionId"></param>
        /// <returns></returns>
        [NotNull]
        public static XElement SetTableStyles([NotNull] this XElement source, int revisionId)
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
                XElement oldTcPr = cell.Element(W + "tcPr");
                XElement newTcPr = new XElement(W + "tcPr");

                // Preserve the cell width, if set explicitly.
                if (oldTcPr?.Element(W + "tcW") is XElement width)
                {
                    newTcPr.Add(
                        new XElement(
                            width.Name,
                            width.Attributes(),
                            width.Nodes()));
                }

                // Preserve the cell span, if set explicitly.
                if (oldTcPr?.Element(W + "gridSpan") is XElement span)
                {
                    newTcPr.Add(
                        new XElement(
                            span.Name,
                            span.Attributes(),
                            span.Nodes()));
                }

                // Preserve the cell alignment, if set explicitly.
                if (oldTcPr?.Element(W + "vAlign") is XElement align)
                {
                    newTcPr.Add(
                        new XElement(
                            align.Name,
                            align.Attributes(),
                            align.Nodes()));
                }

                int old =
                    oldTcPr?.Element(W + "tcMar")?.Element(W + "start")?.Attribute(W + "w") is XAttribute oldStart
                        ? (int) oldStart
                        : 0;

                int count =
                    cell.Value.StartsWith("@>")
                        ? cell.Value.Skip(1).TakeWhile(x => x == '>').Count()
                        : 0;

                if (old + count > 0)
                {
                    newTcPr.Add(
                        new XElement(W + "tcMar",
                            new XElement(W + "start",
                                new XAttribute(W + "w", old + count * 144))));
                }

                oldTcPr?.Remove();

                if (!newTcPr.HasElements)
                    continue;

                cell.AddFirst(newTcPr);

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