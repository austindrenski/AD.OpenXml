using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml.Standard;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Html
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class ConvertTablesExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertTables(this XElement element)
        {
            IEnumerable<XElement> tables = element.Descendants("tbl").ToArray();

            IEnumerable<XElement> items = 
                tables.Select(x => x.Elements("tr").FirstOrDefault())
                      .ToArray();

            foreach (XElement item in items.Elements("td"))
            {
                XElement cell = new XElement("th", item);
                cell.Elements().Promote();
                item.AddAfterSelf(cell);
            }

            tables.Elements("tr").FirstOrDefault()?.Elements("td").Remove();
            tables.Elements("tblPr").Remove();
            tables.Elements("tblGrid").Remove();

            foreach (XElement item in tables)
            {
                item.AddAfterSelf(new XElement("table", item.Elements()));
            }

            tables.Remove();

            return element;
        }
    }
}
