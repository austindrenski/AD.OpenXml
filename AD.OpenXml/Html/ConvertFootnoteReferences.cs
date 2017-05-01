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
    public static class ConvertFootnoteReferencesExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement ConvertFootnoteReferences(this XElement element)
        {
            IEnumerable<XElement> references = element.Descendants("footnoteReference").ToArray();
            foreach (XElement reference in references)
            {
                XElement link = new XElement("a", reference.Attribute("id")?.Value);
                link.SetAttributeValue("href", "#" + reference.Attribute("id")?.Value);
                XElement sup = new XElement("sup", link);
                reference.AddAfterSelf(sup);
                reference.Remove();
            }
            element.Descendants("rPr").Where(x => x.Descendants("rStyle").Any(y => y.Attribute("val")?.Value == "FootnoteReference")).Remove();
            return element;
        }
    }
}
