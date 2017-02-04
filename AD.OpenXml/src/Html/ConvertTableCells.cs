using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AjdExtensions.Html
{
    [PublicAPI]
    public static class ConvertTableCellsExtensions
    {
        public static XElement ConvertTableCells(this XElement element)
        {
            IEnumerable<XElement> items = element.Descendants("tc").ToArray();

            foreach (XElement item in items)
            {
                XElement cell = new XElement("td", item.Descendants("p"));
                if (item.Descendants("jc").Any())
                {
                    cell.SetAttributeValue("class", item.Descendants("jc").Attributes("val").Select(x => x.Value).Concat());
                }
                if (cell.Elements("p").Any())
                {
                    cell.Elements().Promote();
                }
                item.AddAfterSelf(cell);
                item.Remove();
            }

            element.Descendants("jc").Remove();

            return element;
        }
    }
}
