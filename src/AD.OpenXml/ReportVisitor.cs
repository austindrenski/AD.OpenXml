using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <inheritdoc />
    /// <summary>
    /// Defines a visitor to create human-readable OpenXML content.
    /// </summary>
    [PublicAPI]
    public class ReportVisitor : OpenXmlVisitor
    {
        #region Fields

        /// <summary>
        /// The attributes that are supported.
        /// </summary>
        protected ISet<XName> SupportedAttributes =
            new HashSet<XName>
            {
                // TODO: These are found inside of w:drawing. Handle explicitly there.
                "cx",
                "cy",
                "distT",
                "distB",
                "distL",
                "distR",
                "t",
                "b",
                "l",
                "r",
                "id",
                "name",
                "uri",
                "x",
                "y",
                "prst",
                R + "embed",
                R + "id",
                W + "firstColumn",
                W + "firstRow",
                W + "fldCharType",
                W + "id",
                W + "lastColumn",
                W + "lastRow",
                W + "left",
                W + "noHBand",
                W + "noVBand",
                W + "pos",
                W + "right",
                W + "type",
                W + "val",
                W + "w",
                Xml + "space"
            };

        /// <summary>
        /// The elements that are supported.
        /// </summary>
        protected ISet<XName> UnsupportedElements =
            new HashSet<XName>
            {
                W + "bCs",
                W + "bookmarkEnd",
                W + "bookmarkStart",
                W + "color",
                W + "hideMark",
                W + "iCs",
                W + "keepNext",
                W + "lang",
                W + "lastRenderedPageBreak",
                W + "noProof",
                W + "noWrap",
                W + "numPr",
                W + "proofErr",
                W + "rFonts",
                W + "spacing",
                W + "sz",
                W + "szCs",
                W + "tblPrEx",
            };

        /// <summary>
        /// The attributes that are supported.
        /// </summary>
        protected ISet<string> SupportedStyles =
            new HashSet<string>
            {
                "Caption",
                "CaptionFigure",
                "CaptionTable",
                "CommentReference",
                "Emphasis",
                "FiguresTablesSourceNote",
                "FootnoteReference",
                "Heading1",
                "Heading2",
                "Heading3",
                "Heading4",
                "Heading5",
                "Heading6",
                "Heading7",
                "Heading8",
                "Heading9",
                "ListParagraph",
                "Strong"
            };

        #endregion

        #region Constructors

        /// <inheritdoc />
        public ReportVisitor() : base(true) {}

        #endregion

        #region Visits

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAlternateContent(XElement alternate)
        {
            XElement contentControl =
                alternate.Element(MC + "Choice")?
                         .Descendants(W + "sdt")
                         .FirstOrDefault();

            return
                contentControl != null
                    ? Visit(contentControl)
                    : MakeLiftable(alternate.Element(MC + "Choice"));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAttribute(XAttribute attribute)
            => attribute.IsNamespaceDeclaration || SupportedAttributes.Contains(attribute.Name)
                   ? base.VisitAttribute(attribute)
                   : null;

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitBreak(XElement br)
            => "page" == (string) br.Attribute(W + "type")
                   ? null
                   : base.VisitBreak(br);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitElement(XElement element)
            => UnsupportedElements.Contains(element.Name)
                   ? null
                   : base.VisitElement(element);

        /// <inheritdoc />
        [Pure]
        [NotNull]
        protected override XObject VisitHyperlink(XElement hyperlink)
            => new XElement(W + "hyperlink",
                hyperlink.Attribute(R + "id"),
                new XElement(W + "r",
                    new XElement(W + "rPr",
                        new XElement(W + "rStyle",
                            new XAttribute(W + "val", "Hyperlink")),
                        hyperlink.Descendants(W + "i").Any() ||
                        hyperlink.Descendants().Attributes(W + "val").Any(x => (string) x == "Emphasis")
                            ? new XElement(W + "i")
                            : null),
                    new XElement(W + "t", hyperlink.Value)));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitOLEObject(XElement oleObject) => null;

        /// <inheritdoc />
        [Pure]
        [NotNull]
        protected override XObject VisitParagraphProperties(XElement properties)
        {
            XNode[] nodes =
                properties.Nodes()
                          .Where(x => !(x is XElement e && e.Name == W + "rPr"))
                          .Where(x => !(x is XElement e && e.Name == W + "rStyle"))
                          .ToArray();

            return
                new XElement(
                    properties.Name,
                    properties.Attributes(),
                    nodes);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitParagraphStyle(XElement style)
            => SupportedStyles.Contains((string) style.Attribute(W + "val"))
                   ? base.VisitParagraphStyle(style)
                   : null;

        /// <inheritdoc />
        [Pure]
        [NotNull]
        protected override XObject VisitSectionProperties(XElement section)
            => new XElement(W + "sectPr",
                section.Element(W + "cols"),
                section.Element(W + "docGrid"),
                section.Element(W + "pgMar"),
                section.Element(W + "pgSz"));

        #endregion
    }
}