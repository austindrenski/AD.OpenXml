using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AjdExtensions.Html
{
    [PublicAPI]
    public static class ConvertTableCaptionsExtensions
    {
        public static XElement ConvertTableCaptions(this XElement element)
        {
            IList<XElement> tables = element.Descendants("table").ToArray();
            IList<XElement> captions = tables.Select(x => x.Previous()).ToArray();

            for (int i = 0; i < tables.Count; i++)
            {
                captions[i].Remove();
                XElement caption = new XElement("caption", captions[i]);
                caption.Elements().Promote();
                tables[i].AddFirst(caption);
            }
            
            return element;
        }
    }
}
