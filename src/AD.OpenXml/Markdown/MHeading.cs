﻿using System;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

// TODO: temporary until StringSegment is a readonly struct in 2.1
// ReSharper disable ImpureMethodCallOnReadonlyValueField

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
            if (!Accept(in text))
            {
                throw new ArgumentException($"Heading must begin with 1-6 '#' characters followed by a ' ' character: '{Text}'");
            }

            StringSegment t = Normalize(in text);

            Level = t.IndexOf(' ');
            Heading = t.Subsegment(Level + 1);
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
            StringSegment trimmed = Normalize(in segment);

            return
                trimmed.Length > 2 &&
                trimmed.StartsWith("# ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("## ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("### ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("#### ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("##### ", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("###### ", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Normalizes the segment by trimming (in order):
        ///   1) up to three ' ' characters from the start;
        ///   2) all ' ' from the end;
        ///   3) all '#' from the end;
        ///   4) one ' ' from the end;
        ///   5) normalizing inner whitespace.
        /// </summary>
        /// <param name="segment">
        /// The segment to normalize.
        /// </param>
        /// <returns>
        /// The normalized segment.
        /// </returns>
        private static StringSegment Normalize(in StringSegment segment)
        {
            return segment.TrimStart(' ', 3).TrimEnd().TrimEnd('#').TrimEnd(' ', 1).NormalizeInner(' ');
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString()
        {
            return $"{new string('#', Level)} {Heading}";
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