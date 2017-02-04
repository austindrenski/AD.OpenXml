using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Html
{
    [PublicAPI]
    public static class ConvertSimpleExtensions
    {
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