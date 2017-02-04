using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AjdExtensions.Html
{
    [PublicAPI]
    public static class ConvertBoldRunsExtensions
    {
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