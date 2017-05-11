using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace AD.OpenXml.Core.Elements
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class MergeRunsExtensions
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static XElement MergeRuns(this XElement element)
        {
            IEnumerable<XElement> paragraphs = element.Descendants(W + "p").ToArray();
            foreach (XElement paragraph in paragraphs)
            {
                IEnumerable<XElement> runs = paragraph.Elements(W + "r").ToArray();
                foreach (XElement run in runs)
                {
                    if (run.Element(W + "t") is null)
                    {
                        continue;
                    }
                    if (run.Element(W + "rPr")?.ToString() != run.Next()?.Element(W + "rPr")?.ToString())
                    {
                        continue;
                    }
                    if (run.Next()?.Name != W + "r")
                    {
                        continue;
                    }
                    if (run.Element(W + "fldChar") != null)
                    {
                        continue;
                    }
                    if (run.Next()?.Element(W + "fldChar") != null)
                    {
                        continue;
                    }
                    if (!run.Next()?.Elements(W + "t").Any() ?? false)
                    {
                        run.Next()?.Add(new XElement(W + "t"));
                    }
                    XElement xElement = run.Next()?.Element(W + "t");
                    if (xElement != null)
                    {
                        xElement.Value = run.Value + xElement.Value;

                        xElement.Value = xElement.Value.Replace("  ", null);

                        if (xElement.Value.Length != xElement.Value.Trim().Length)
                        {
                            xElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                        }
                    }
                    run.Remove();
                }
            }
            return element;
        }
    }
}
