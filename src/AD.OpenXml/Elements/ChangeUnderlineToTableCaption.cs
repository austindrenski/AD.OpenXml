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
    public static class ChangeUnderlineToTableCaptionExtensions
    {
        [NotNull] static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;
        [NotNull] static readonly XNamespace Xml = XNamespace.Xml;

        /// <summary>
        /// Removes all &lt;u [val=...]/&gt; descendant elements from the &lt;rPr [...]/&gt; elements
        /// and places a &lt;pStyle val="CaptionTable" /&gt; on the &lt;pPr [...]/&gt; elements.
        ///
        /// This method works on the existing <see cref="XElement"/> and returns a reference to it for a fluent syntax.
        /// </summary>
        /// <param name="element">The element to search for descendants.</param>
        /// <returns>A reference to the existing <see cref="XElement"/>. This is returned for use with fluent syntax calls.</returns>
        /// <exception cref="System.ArgumentException"/>
        /// <exception cref="System.ArgumentNullException"/>
        [NotNull]
        public static XElement ChangeUnderlineToTableCaption([NotNull] this XElement element)
        {
            XElement[] paragraphs =
                element.Descendants(W + "u")
                       .Select(x => x.Parent)
                       .Where(x => x?.Name == W + "rPr")
                       .Select(x => x.Parent)
                       .Where(x => x?.Name == W + "r")
                       .Select(x => x.Parent)
                       .Where(x => x?.Name == W + "p")
                       .Where(x => x.Next()?.Name == W + "tbl" || (x.Next()?.Value.Contains('{') ?? false))
                       .Distinct()
                       .ToArray();

            foreach (XElement item in paragraphs)
            {
                item.ReplaceWith(item.AddTableCaption());
            }

            return element;
        }

        [NotNull]
        static XNode Triage(XNode node)
        {
            if (!(node is XElement e))
                return node;

            if (e.Name != W + "p")
                return node;

            if (!(e.NextNode is XElement next))
                return node;

            if (!(next.Name != W + "tbl"))
                return node;

            if (!(e.Element(W + "r") is XElement r))
                return node;

            if (!(r.Element(W + "rPr") is XElement rPr))
                return node;

            return rPr.Element(W + "u") is null ? node : e.AddTableCaption();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="element"></param>
        [Pure]
        [NotNull]
        static XElement AddTableCaption([NotNull] this XElement element)
        {
            string style = element.Value.Contains("[APPENDIX]") ? "9" : "1";

            return
                new XElement(
                    element.Name,
                    element.Attributes(),
                    new XElement(W + "pPr",
                        new XElement(W + "pStyle",
                            new XAttribute(W + "val", "CaptionTable"))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val",
                                    "Strong"))),
                        new XElement(W + "t",
                            new XAttribute(Xml + "space", "preserve"),
                            new XText("Table "))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "fldChar",
                            new XAttribute(W + "fldCharType", "begin"))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "instrText",
                            new XAttribute(Xml + "space", "preserve"),
                            new XText($" STYLEREF {style} \\s "))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "fldChar",
                            new XAttribute(W + "fldCharType", "separate"))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "t", "0")),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "fldChar",
                            new XAttribute(W + "fldCharType", "end"))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "t", ".")),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "fldChar",
                            new XAttribute(W + "fldCharType", "begin"))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "instrText",
                            new XAttribute(Xml + "space", "preserve"),
                            new XText($" SEQ Table \\* ARABIC \\s {style} "))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "fldChar",
                            new XAttribute(W + "fldCharType", "separate"))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "t", "0")),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "fldChar",
                            new XAttribute(W + "fldCharType", "end"))),
                    new XElement(W + "r",
                        new XElement(W + "rPr",
                            new XElement(W + "rStyle",
                                new XAttribute(W + "val", "Strong"))),
                        new XElement(W + "t",
                            new XAttribute(Xml + "space", "preserve"),
                            new XText(" "))),
                    element.Nodes().Select(RemoveAppendixIdentifier));
        }

        [CanBeNull]
        static XNode RemoveAppendixIdentifier(XNode node)
        {
            if (!(node is XElement e))
                return node;

            if (e.Name == W + "u")
                return null;

            foreach (XElement text in e.Descendants(W + "t"))
            {
                text.Value = text.Value.Replace("[", null);
                text.Value = text.Value.Replace("APPENDIX", null);
                text.Value = text.Value.Replace("]", null);
            }

            return e;
        }
    }
}