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
        /// <summary>
        /// The attributes that may be returned.
        /// </summary>
        protected ISet<XName> SupportedAttributes { get; }

        /// <summary>
        /// The elements that may be returned.
        /// </summary>
        protected ISet<XName> SupportedElements { get; }

        /// <inheritdoc />
        public ReportVisitor() : base(true)
        {
        }

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



//// Remove editing attributes.
//.RemoveRsidAttributes()
//
//// Remove elements that should never exist in-line.
//.RemoveByAll(W + "bCs")
//.RemoveByAll(W + "bookmarkEnd")
//.RemoveByAll(W + "bookmarkStart")
//.RemoveByAll(W + "color")
//.RemoveByAll(W + "hideMark")
//.RemoveByAll(W + "iCs")
//.RemoveByAll(W + "keepNext")
//.RemoveByAll(W + "lang")
//.RemoveByAll(W + "lastRenderedPageBreak")
//.RemoveByAll(W + "noProof")
//.RemoveByAll(W + "noWrap")
//.RemoveByAll(W + "numPr")
//.RemoveByAll(W + "proofErr")
//.RemoveByAll(W + "rFonts")
//.RemoveByAll(W + "spacing")
//.RemoveByAll(W + "sz")
//.RemoveByAll(W + "szCs")
//.RemoveByAll(W + "tblPrEx")
    }
}