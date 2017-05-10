using System;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Standard.Properties;
using AD.OpenXml.Standard.Visitors;
using JetBrains.Annotations;

namespace AD.OpenXml.Standard.Visits
{
    /// <summary>
    /// 
    /// </summary>
    public sealed class StyleVisit : IOpenXmlVisit
    {
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        public IOpenXmlVisitor Result { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        public StyleVisit(IOpenXmlVisitor subject)
        {
            XElement styles = Execute(subject.Styles.Clone());

            Result =
                new OpenXmlVisitor(
                    subject.ContentTypes,
                    subject.Document,
                    subject.DocumentRelations,
                    subject.Footnotes,
                    subject.FootnoteRelations,
                    styles,
                    subject.Numbering,
                    subject.Charts);
        }

        [Pure]
        [NotNull]
        private static XElement Execute([NotNull] XElement styles)
        {
            if (styles is null)
            {
                throw new ArgumentNullException(nameof(styles));
            }

            XElement results =
                new XElement(
                    styles.Name,
                    styles.Attributes(),
                    XElement.Parse(Resources.DocDefaults),
                    XElement.Parse(Resources.Normal),
                    XElement.Parse(Resources.Appendix),
                    XElement.Parse(Resources.Bibliography),
                    XElement.Parse(Resources.Box),
                    XElement.Parse(Resources.BoxCaption),
                    XElement.Parse(Resources.BoxListBullet),
                    XElement.Parse(Resources.BoxSourceNote),
                    XElement.Parse(Resources.BlueTableBasic),
                    XElement.Parse(Resources.CaptionFigure),
                    XElement.Parse(Resources.CaptionTable),
                    XElement.Parse(Resources.Emphasis),
                    XElement.Parse(Resources.ExecutiveSummary1stParagraph),
                    XElement.Parse(Resources.ExecutiveSummaryHighlights),
                    XElement.Parse(Resources.ExecutiveSummarySidebars),
                    XElement.Parse(Resources.FigureTableSourceNote),
                    XElement.Parse(Resources.FootnoteReference),
                    XElement.Parse(Resources.FootnoteText),
                    XElement.Parse(Resources.Heading1),
                    XElement.Parse(Resources.Heading2),
                    XElement.Parse(Resources.Heading3),
                    XElement.Parse(Resources.Heading4),
                    XElement.Parse(Resources.Heading5),
                    XElement.Parse(Resources.Heading6),
                    XElement.Parse(Resources.Heading7),
                    XElement.Parse(Resources.Heading8),
                    XElement.Parse(Resources.Heading9),
                    XElement.Parse(Resources.Hyperlink),
                    XElement.Parse(Resources.ListBullet),
                    XElement.Parse(Resources.PreHeading),
                    XElement.Parse(Resources.Strong),
                    XElement.Parse(Resources.StyleNotImplemented),
                    XElement.Parse(Resources.Subscript),
                    XElement.Parse(Resources.Superscript),
                    XElement.Parse(Resources.TableOfFigures),
                    XElement.Parse(Resources.TOC1),
                    XElement.Parse(Resources.TOC2),
                    XElement.Parse(Resources.TOC3),
                    XElement.Parse(Resources.TOC4),
                    XElement.Parse(Resources.TOCHeading));

            foreach (XElement style in results.Elements())
            {
                if (style.Elements().First().Name == W + "name")
                {
                    continue;
                }

                XElement name = style.Element(W + "name");
                name?.Remove();
                style.AddFirst(name);
            }

            foreach (XElement style in results.Elements())
            {
                if (style.Element(W + "name")?.Next()?.Name == W + "next")
                {
                    continue;
                }

                XElement next = style.Element(W + "next");
                next?.Remove();
                style.Element(W + "name")?.AddAfterSelf(next);
            }


            return results;
        }
    }
}