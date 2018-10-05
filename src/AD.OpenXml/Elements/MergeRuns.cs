using System;
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
        [NotNull] static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] static readonly Regex Spaces = new Regex(@"(\s{2,})", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        ///
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        [NotNull]
        public static XElement MergeRuns([NotNull] this XElement element)
        {
            IEnumerable<XElement> paragraphs = element.Descendants(W + "p").ToArray();
            foreach (XElement paragraph in paragraphs)
            {
                ProcessRuns(paragraph);

                // TODO: this fixes a problem with leading spaces showing up at the start of paragraphs.
                if (paragraph.Elements(W + "r").FirstOrDefault() is XElement first)
                {
                    if (first.Element(W + "t") is XElement text)
                    {
                        if (text.Value.StartsWith(" "))
                            text.Value = text.Value.TrimStart();
                    }
                }

                // ReSharper disable once InvertIf
                if (paragraph.Elements(W + "r").LastOrDefault() is XElement last)
                {
                    // ReSharper disable once InvertIf
                    if (last.Element(W + "t") is XElement text)
                    {
                        if (text.Value.EndsWith(" "))
                            text.Value = text.Value.TrimEnd();
                    }
                }
            }

            return element;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="text">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public static void RemoveDuplicateSpacing([NotNull] XElement text)
        {
            if (text is null)
                throw new ArgumentNullException(nameof(text));

            text.Value = Spaces.Replace((string) text, " ");

            if (text.Value.StartsWith(" ") || text.Value.EndsWith(" "))
                text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        }

        static void ProcessRuns([NotNull] XElement paragraph)
        {
            IEnumerable<XElement> runs = paragraph.Elements(W + "r").ToArray();
            foreach (XElement run in runs)
            {
                XElement currentText = run.Element(W + "t");

                if (currentText is null)
                    continue;

                RemoveDuplicateSpacing(currentText);

                XElement currentRpr = run.Element(W + "rPr");
                XElement nextRpr = run.Next()?.Element(W + "rPr");

                // TODO: This is a weak heuristic. Handle child node comparison explicitly.
                if (currentRpr?.Elements().Count() != nextRpr?.Elements().Count())
                    continue;

                XElement currentRStyle = currentRpr?.Element(W + "rStyle");
                XElement nextRStyle = nextRpr?.Element(W + "rStyle");

                if ((string) currentRStyle?.Attribute(W + "val") != (string) nextRStyle?.Attribute(W + "val"))
                    continue;

                if (run.Next()?.Name != W + "r")
                    continue;

                if (run.Element(W + "drawing") != null)
                    continue;

                if (run.Element(W + "fldChar") != null)
                    continue;

                if (run.Next()?.Element(W + "fldChar") != null)
                    continue;

                if (run.Next()?.Element(W + "footnoteReference") != null)
                    continue;

                if (!run.Next()?.Elements(W + "t").Any() ?? false)
                    run.Next()?.Add(new XElement(W + "t"));

                if (run.Next()?.Element(W + "t") is XElement nextText)
                {
                    nextText.Value = run.Value + nextText.Value;
                    RemoveDuplicateSpacing(nextText);
                }

                run.Remove();
            }
        }
    }
}