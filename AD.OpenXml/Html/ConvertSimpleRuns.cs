using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Html
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class ConvertSimpleExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertSimpleRuns(this XElement element)
        {
            IEnumerable<XElement> items =
                element.Descendants("r")
                       .ToArray();

            IEnumerable<XElement> simpleRuns =
                items.Where(x => !x.Descendants("b").Any()
                              && !x.Descendants("i").Any()
                              && x.Descendants("rStyle").Attributes("val").All(y => y.Value != "Strong" && y.Value != "Emphasis"));

            simpleRuns.Promote();
            return element;
        }
    }
}