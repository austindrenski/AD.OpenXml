using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Html
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class ConvertBoldRunsExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertBoldRuns(this XElement element)
        {
            IEnumerable<XElement> items =
                element.Descendants("r")
                       .ToArray();

            IEnumerable<XElement> boldRuns = 
                items.Where(x => x.Descendants("b").Any() 
                              || x.Descendants("rStyle").Attributes("val").Any(y => y.Value == "Strong"));

            foreach (XElement item in boldRuns)
            {
                item.AddAfterSelf(new XElement("strong", item.Value));
                item.Remove();
            }

            return element;
        }
    }
}