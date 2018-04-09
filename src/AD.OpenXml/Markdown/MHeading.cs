﻿using System;
using System.Xml.Linq;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;
using AD.ApiExtensions.Primitives;

// TODO: temporary until StringSegment is a readonly struct in 2.1
// ReSharper disable ImpureMethodCallOnReadonlyValueField

namespace AD.OpenXml.Markdown
{
    /// <inheritdoc cref="MNode"/>
    /// <inheritdoc cref="IEquatable{T}"/>
    /// <summary>
    /// Represents a Markdown header node.
    /// </summary>
    /// <remarks>
    /// Note: closing sequences are not supported.
    /// See: http://spec.commonmark.org/0.28/#atx-headings
    /// </remarks>
    [PublicAPI]
    public class MHeading : MNode, IEquatable<MHeading>
    {
        /// <summary>
        /// The text of the heading.
        /// </summary>
        [NotNull]
        public MText Heading { get; }

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
        public MHeading(in StringSegment text)
        {
            if (!Accept(in text))
            {
                throw new ArgumentException($"Heading must begin with 1-6 '#' characters followed by a ' ' character: '{text}'");
            }

            StringSegment normalized = Normalize(in text);
            Level = normalized.IndexOf(' ');
            Heading = normalized.Subsegment(Level + 1);
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
        [Pure]
        private static StringSegment Normalize(in StringSegment segment)
        {
            return segment.TrimStart(' ', 3).TrimEnd().TrimEnd('#').TrimEnd(' ', 1).NormalizeInner(' ');
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"{new string('#', Level)} {Heading}";
        }

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml()
        {
            return new XElement($"h{Level}", Heading.ToHtml());
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
                    new XElement(W + "r", Heading.ToOpenXml()));
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals([CanBeNull] MHeading other)
        {
            return !(other is null) && Heading.Equals(other.Heading) && Level == other.Level;
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is MHeading heading && Equals(heading);
        }

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            unchecked
            {
                return (397 * Heading.GetHashCode()) ^ Level.GetHashCode();
            }
        }

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Markdown.MHeading" /> objects are equal.
        /// </summary>
        /// <param name="left">
        /// The first value to compare.
        /// </param>
        /// <param name="right">
        /// The second value to compare.
        /// </param>
        /// <returns>
        /// true if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==([CanBeNull] MHeading left, [CanBeNull] MHeading right)
        {
            return !(left is null) && !(right is null) && left.Equals(right);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Markdown.MHeading" /> objects have different values.
        /// </summary>
        /// <param name="left">
        /// The first value to compare.
        /// </param>
        /// <param name="right">
        /// The second value to compare.
        /// </param>
        /// <returns>
        /// true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=([CanBeNull] MHeading left, [CanBeNull] MHeading right)
        {
            return left is null || right is null || !left.Equals(right);
        }
    }
}