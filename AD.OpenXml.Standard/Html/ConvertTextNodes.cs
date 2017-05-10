using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Html
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class ConvertTextNodesExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertTextNodes(this XElement element)
        {
            element.Descendants("t").Promote();
            return element;
        }
    }
}
