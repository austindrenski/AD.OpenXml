using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// Extension methods to replace &lt;u/&gt; elements with &lt;pStyle val=[...]/&gt; elements.
    /// </summary>
    [PublicAPI]
    public static class ChangeUnderlineToSourceNoteExtensions
    {
        [NotNull] static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Removes all &lt;u [val=...]/&gt; descendant elements from the &lt;rPr [...]/&gt; elements
        /// and places a &lt;pStyle val="FiguresTablesSourceNote" /&gt; on the &lt;pPr [...]/&gt; elements.
        ///
        /// This method works on the existing <see cref="XElement"/> and returns a reference to it for a fluent syntax.
        /// </summary>
        /// <param name="element">The element to search for descendants.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        [NotNull]
        public static XElement ChangeUnderlineToSourceNote([NotNull] this XElement element)
        {
            IEnumerable<XElement> paragraphs =
                element.Descendants(W + "u")
                       .Select(x => x.Parent)
                       .Where(x => x?.Name == W + "rPr")
                       .Select(x => x.Parent)
                       .Where(x => x?.Name == W + "r")
                       .Select(x => x.Parent)
                       .Where(x => x?.Name == W + "p")
                       .Where(x => (x.Previous()?.DescendantsAndSelf(W + "tbl").Any() ?? false)
                                || (x.Previous()?.DescendantsAndSelf(W + "drawing").Any() ?? false)
                                || (x.Previous()?.DescendantsAndSelf(W + "u").Any() ?? false)
                                || (x.Previous()?.Value.Contains('{') ?? false))
                       .Distinct()
                       .ToArray();

            foreach (XElement item in paragraphs)
            {
                item.Descendants(W + "pStyle").Remove();
                if (!item.Elements(W + "pPr").Any())
                    item.AddFirst(new XElement(W + "pPr"));
                item.Element(W + "pPr")?.AddFirst(new XElement(W + "pStyle", new XAttribute(W + "val", "FiguresTablesSourceNote")));
            }

            return element;
        }
    }
}