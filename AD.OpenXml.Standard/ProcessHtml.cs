using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Standard.Html;
using AD.Xml.Standard;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard
{
    /// <summary>
    /// Performs a variety of processing to transform the XML node into a well-formed HTML document.
    /// </summary>
    [PublicAPI]
    public static class ProcessHtmlExtensions
    {
        /// <summary>
        /// Returns an XML element as a well-formed HTML document  of the form &lt;html&gt;&lt;head&gt;...&lt;/head&gt;&lt;body&gt;...&lt;/body&gt;.
        /// </summary>
        /// <param name="element">The XML node to process.</param>
        /// <returns>A well-formed HTML document.</returns>
        public static XElement ProcessHtml(this XElement element)
        {
            return element.RemoveAttributesBy(XNamespace.Xml + "space")
                          .RemoveNamespaces()
                          .Element("body")
                          .BodyToHtml("Word to HTML Testing", "..\\..\\USITC332.css")
                          .AddToAll(
                               x => x.Name == "head",
                              "script",
                              new XAttribute("type", "text/javascript"),
                              new XAttribute("src", "https://cdn.mathjax.org/mathjax/latest/MathJax.js?config=MML_CHTML"),
                              " ")
                          .ConvertFootnoteReferences()
                          .ConvertTableCells()
                          .ConvertTables()
                          .ConvertTableCaptions()
                          .ConvertHeadings()
                          .ConvertBoldRuns()
                          .ConvertItalicRuns()
                          .ConvertStyleToClass()
                          .ConvertSimpleRuns()
                          .ConvertTextNodes()
                          .RemoveByAllIfEmpty("rPr")
                          .RemoveByAllIfEmpty("pPr")
                          .RemoveByAll("sectPr")
                          .ChangeXAttributeValues("class", "CaptionTable", "chapter");
        }
    }
}
