using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class MergeRunsExtensions
    {
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] private static readonly Regex Spaces = new Regex(@"(\s{2,})", RegexOptions.Compiled | RegexOptions.IgnoreCase);

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
                ProcessRuns(paragraph);

                if (paragraph.Elements(W + "r").FirstOrDefault() is XElement first)
                {
                    if (first.Element(W + "t") is XElement text)
                    {
                        if (text.Value.StartsWith(" "))
                        {
                            text.Value = text.Value.TrimStart();
                        }
                    }
                }

                if (paragraph.Elements(W + "r").LastOrDefault() is XElement last)
                {
                    if (last.Element(W + "t") is XElement text)
                    {
                        if (text.Value.EndsWith(" "))
                        {
                            text.Value = text.Value.TrimEnd();
                        }
                    }
                }
            }

            return element;
        }

        private static void ProcessRuns(XElement paragraph)
        {
            IEnumerable<XElement> runs = paragraph.Elements(W + "r").ToArray();
            foreach (XElement run in runs)
            {
                XElement currentText = run.Element(W + "t");

                if (currentText is null)
                {
                    continue;
                }

                RemoveDuplicateSpacing(currentText);

                if ((string) run.Element(W + "rPr") != (string) run.Next()?.Element(W + "rPr"))
                {
                    continue;
                }

                if (run.Next()?.Name != W + "r")
                {
                    continue;
                }

                if (run.Element(W + "drawing") != null)
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

                if (run.Next()?.Element(W + "t") is XElement nextText)
                {
                    nextText.Value = $"{run} {nextText}";
                    RemoveDuplicateSpacing(nextText);
                }

                run.Remove();
            }
        }

        private static void RemoveDuplicateSpacing(XElement text)
        {
            text.Value = Spaces.Replace((string) text, " ");

            if (text.Value.StartsWith(" ") || text.Value.EndsWith(" "))
            {
                text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            }
        }
    }
}