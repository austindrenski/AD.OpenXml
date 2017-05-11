using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml.Standard;
using JetBrains.Annotations;

namespace AD.OpenXml.Html
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class ToHeadingsExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertHeadings(this XElement element)
        {
            return element.ConvertHeadings("1")
                          .ConvertHeadings("2")
                          .ConvertHeadings("3")
                          .ConvertHeadings("4")
                          .ConvertHeadings("5")
                          .ConvertHeadings("6")
                          .ConvertHeadings("7");
        }

        private static XElement ConvertHeadings(this XElement element, string number)
        {
            IEnumerable<XElement> items =
                element.DescendantsAndSelf("pStyle")
                       .Where(x => x.Attribute("val")?.Value == "Heading" + number)
                       .Select(x => x.Parent)
                       .Select(x => x?.Parent)
                       .ToArray();
            foreach (XElement item in items)
            {
                XElement heading =
                    new XElement("h" + number, item?.Elements()
                                           .Where(x => x.Name != "pPr")
                                           .Detach()
                                           .Select(x => x.Value)
                                           .Concat()
                                           .Replace("  ", " "));
                item?.AddAfterSelf(heading);
                item?.Remove();
            }
            return element;
        }
    }
}