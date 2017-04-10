using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    [PublicAPI]
    public static class SetTableStylesExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        public static XElement SetTableStyles(this XElement element)
        {
            foreach (XElement item in element.Descendants(W + "tblPr"))
            {
                item.RemoveAll();
                item.Add(
                    new XElement(W + "tblStyle",
                        new XAttribute(W + "val", "BlueTableBasic")),
                    new XElement(W + "tblW",
                        new XAttribute(W + "type", "dxa"),
                        new XAttribute(W + "w", "9360")),
                    new XElement(W + "tblLayout",
                        new XAttribute(W + "type", "autofit")),
                    new XElement(W + "tblLook",
                        new XAttribute(W + "val", "04A0"),
                        new XAttribute(W + "firstRow", "1"),
                        new XAttribute(W + "lastRow", "0"),
                        new XAttribute(W + "firstColumn", "0"),
                        new XAttribute(W + "lastColumn", "0"),
                        new XAttribute(W + "noHBand", "0"),
                        new XAttribute(W + "noVBand", "1")));
            }
            element.Descendants(W + "tcPr").Descendants().Where(x => x.Name != W + "vAlign").Remove();
            element.Descendants(W + "trPr").Remove();
            element.Descendants(W + "gridCol").Attributes(W + "w").Remove();
            
            return element;
        }
    }
}