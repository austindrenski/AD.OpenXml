using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AjdExtensions.Html
{
    [PublicAPI]
    public static class ConvertTextNodesExtensions
    {
        public static XElement ConvertTextNodes(this XElement element)
        {
            element.Descendants("t").Promote();
            return element;
        }
    }
}
