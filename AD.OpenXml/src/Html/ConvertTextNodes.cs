using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Html
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
