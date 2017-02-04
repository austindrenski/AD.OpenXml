using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AjdExtensions.Html
{
    [PublicAPI]
    public static class ConvertItalicRunsExtensions
    {
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