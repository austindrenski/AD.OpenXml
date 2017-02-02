using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    [PublicAPI]
    public static class MergeRunsExtensions
    {
        private static XNamespace _w = XNamespaces.OpenXmlWordprocessingmlMain;

        public static XElement MergeRuns(this XElement element)
        {
            IEnumerable<XElement> paragraphs = element.Descendants(_w + "p").ToArray();
            foreach (XElement paragraph in paragraphs)
            {
                IEnumerable<XElement> runs = paragraph.Elements(_w + "r").ToArray();
                foreach (XElement run in runs)
                {
                    if (run.Element(_w + "t") == null)
                    {
                        continue;
                    }
                    if (run.Element(_w + "rPr")?.ToString() != run.Next()?.Element(_w + "rPr")?.ToString())
                    {
                        continue;
                    }
                    if (run.Next()?.Name != _w + "r")
                    {
                        continue;
                    }
                    if (!run.Next()?.Elements(_w + "t").Any() ?? false)
                    {
                        run.Next()?.Add(new XElement(_w + "t"));
                    }
                    XElement xElement = run.Next()?.Element(_w + "t");
                    if (xElement != null)
                    {
                        xElement.Value = run.Value + xElement.Value;
                    }
                    run.Remove();
                }
            }
            return element;
        }
    }
}
