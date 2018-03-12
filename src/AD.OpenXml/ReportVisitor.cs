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
        protected ISet<XName> SupportedAttributes { get; }

        /// <summary>
        /// The elements that may be returned.
        /// </summary>
        protected ISet<XName> SupportedElements { get; }

        /// <inheritdoc />
        protected override IDictionary<string, XElement> Charts { get; set; }

        /// <inheritdoc />
        protected override IDictionary<string, (string mime, string description, string base64)> Images { get; set; }

        /// <inheritdoc />
        protected ReportVisitor(bool returnOnDefault) : base(returnOnDefault)
        {
        }
    }
}