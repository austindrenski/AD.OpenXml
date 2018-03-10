using System;
using System.Collections.Generic;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <inheritdoc />
    /// <summary>
    /// Defines a visitor to create human-readable OpenXML content.
    /// </summary>
    public class ReportVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// The attributes that may be returned.
        /// </summary>
        protected override ISet<XName> SupportedAttributes { get; }

        /// <summary>
        /// The elements that may be returned.
        /// </summary>
        protected override ISet<XName> SupportedElements { get; }

        /// <summary>
        /// The mapping between OpenXML names and supported names.
        /// </summary>
        protected override IDictionary<XName, XName> Renames { get; }

        /// <summary>
        /// The mapping of chart id to node.
        /// </summary>
        protected override IDictionary<string, XElement> Charts { get; set; }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitElement(XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            switch (element)
            {
                case XElement e when e.Name.LocalName == "body":
                {
                    return VisitBody(e);
                }
                case XElement e when e.Name.LocalName == "drawing":
                {
                    return VisitDrawing(e);
                }
                case XElement e when e.Name.LocalName == "footnote":
                {
                    return VisitFootnote(e);
                }
                case XElement e when e.Name.LocalName == "p":
                {
                    return VisitParagraph(e);
                }
                case XElement e when e.Name.LocalName == "r":
                {
                    return VisitRun(e);
                }
                case XElement e when e.Name.LocalName == "tbl":
                {
                    return VisitTable(e);
                }
                case XElement e when e.Name.LocalName == "tr":
                {
                    return VisitTableRow(e);
                }
                case XElement e when e.Name.LocalName == "tc":
                {
                    return VisitTableCell(e);
                }
                default:
                {
                    return null;
                }
            }
        }
    }
}