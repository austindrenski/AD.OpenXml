using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace AD.OpenXml.Core.Html
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class ConvertItalicRunsExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertItalicRuns(this XElement element)
        {
            IEnumerable<XElement> items =
                element.Descendants("r")
                       .ToArray();

            IEnumerable<XElement> italicRuns =
                items.Where(x => x.Descendants("i").Any()
                              || x.Descendants("rStyle").Attributes("val").Any(y => y.Value == "Emphasis"));

            foreach (XElement item in italicRuns)
            {
                item.AddAfterSelf(new XElement("em", item.Value));
                item.Remove();
            }

            return element;
        }
    }
}