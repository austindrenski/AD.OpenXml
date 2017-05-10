using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Html
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class ConvertStyleToClassExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertStyleToClass(this XElement element)
        {
            IEnumerable<XElement> items =
                element.Descendants()
                       .Where(x => x.Name == "pStyle" || x.Name == "rStyle")
                       .Where(x => x.Attributes("val").All(y => y.Value != "Strong" && y.Value != "Emphasis"))
                       .ToArray();

            foreach (XElement item in items)
            {
                item.Parent?.Parent?.SetAttributeValue("class", item.Attribute("val")?.Value);
                item.Remove();
            }
            return element;
        }
    }
}