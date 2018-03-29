using System;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

namespace AD.OpenXml.Markdown
{
    /// <inheritdoc />
    /// <summary>
    /// Represents a Markdown header node.
    /// </summary>
    /// <remarks>
    /// Note: closing sequences are not supported.
    /// See: http://spec.commonmark.org/0.28/#atx-headings
    /// </remarks>
    [PublicAPI]
    public class MHeading : MNode
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// The text of the heading.
        /// </summary>
        [NotNull]
        public MNode Heading { get; }

        /// <summary>
        /// The level of the header.
        /// </summary>
        public int Level { get; }

        /// <summary>
        /// Constructs an <see cref="MHeading"/>.
        /// </summary>
        /// <param name="text">
        /// The raw text of the heading.
        /// </param>
        public MHeading(in StringSegment text) : base(in text)
        {
            if (!Accept(in Text))
            {
                throw new ArgumentException($"Heading must begin with 1-6 '#' characters followed by a ' ' character: '{Text}'");
            }

            Level = Text.IndexOf(' ');
            Heading = Text.Subsegment(Level + 1).TrimStart();
        }

        /// <summary>
        /// Checks if the segment is a well-formed Markdown heading.
        /// </summary>
        /// <param name="segment">
        /// The segment to test.
        /// </param>
        /// <returns>
        /// True if the segment is a well-formed Markdown heading; otherwise false.
        /// </returns>
        public static bool Accept(in StringSegment segment)
        {
            StringSegment trimmed = segment.Trim();

            return
                trimmed.Length > 3 &&
                trimmed.StartsWith("# ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("## ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("### ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("#### ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("##### ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("###### ", StringComparison.OrdinalIgnoreCase);
        }

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml()
        {
            return new XElement($"h{Level}", Heading);
        }

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml()
        {
            return
                new XElement(W + "p",
                    new XElement(W + "pPr",
                        new XElement(W + "pStyle",
                            new XAttribute(W + "val", $"Heading{Level}"))),
                    new XElement(W + "r",
                        new XElement(W + "t", Heading)));
        }
    }
}