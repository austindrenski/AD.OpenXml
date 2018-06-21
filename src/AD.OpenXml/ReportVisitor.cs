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
                W + "id",
                W + "type",
                W + "val"
            };

        /// <summary>
        /// The attributes that are not supported.
        /// </summary>
        protected ISet<XName> UnsupportedAttributes =
            new HashSet<XName>
            {
                W + "rsid",
                W + "rsidDel",
                W + "rsidP",
                W + "rsidR",
                W + "rsidRDefault",
                W + "rsidRPr",
                W + "rsidSect",
                W + "rsidTr"
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

        #endregion

        #region Constructors

        /// <inheritdoc />
        public ReportVisitor() : base(true)
        {
        }

        #endregion

        #region Visits

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAttribute(XAttribute attribute)
            => UnsupportedAttributes.Contains(attribute.Name) ? null : base.VisitAttribute(attribute);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitElement(XElement element)
            => UnsupportedElements.Contains(element.Name) ? null : base.VisitElement(element);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitParagraphProperties(XElement properties)
            => new XElement(
                properties.Name,
                properties.Attributes(),
                properties.Nodes().Where(x => !(x is XElement e && e.Name == W + "rPr")));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitOLEObject(XElement oleObject) => null;

        #endregion
    }
}