using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace AD.OpenXml.Core.Html
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class ConvertTableCaptionsExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertTableCaptions(this XElement element)
        {
            IList<XElement> tables = element.Descendants("table").ToArray();
            IList<XElement> captions = tables.Select(x => x.Previous()).Where(x => x != null).ToArray();

            for (int i = 0; i < tables.Count; i++)
            {
                if (i + 1 > captions.Count)
                {
                    continue;
                }

                captions[i].Remove();
                XElement caption = new XElement("caption", captions[i]);
                caption.Elements().Promote();
                tables[i].AddFirst(caption);
            }
            
            return element;
        }
    }
}
