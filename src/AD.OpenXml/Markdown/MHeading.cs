using System;
using System.Xml.Linq;
using JetBrains.Annotations;

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
        /// <param name="text">The raw text of the heading.</param>
        public MHeading(in ReadOnlySpan<char> text)
        {
            if (!Accept(in text))
                throw new ArgumentException($"Heading must begin with 1-6 '#' characters followed by a ' ' character: '{text.ToString()}'");

            ReadOnlySpan<char> normalized = Normalize(in text);
            Level = normalized.IndexOf(' ');
            Heading = normalized.Slice(Level + 1);
        }

        /// <summary>
        /// Checks if the span is a well-formed Markdown heading.
        /// </summary>
        /// <param name="span">The span to test.</param>
        /// <returns>
        /// True if the span is a well-formed Markdown heading; otherwise false.
        /// </returns>
        public static bool Accept(in ReadOnlySpan<char> span)
        {
            if (span.IsEmpty)
                return false;

            int initial = span.IndexOf('#');

            if (initial == -1 || initial > 3)
                return false;

            ReadOnlySpan<char> trimmed = Normalize(in span);

            if (trimmed.Length < 2)
                return false;

            if (trimmed[0] != '#')
                return false;

            for (int i = 0; i < trimmed.Length; i++)
            {
                if (i == 7)
                    return false;

                if (trimmed[i] == '#')
                    continue;

                return trimmed[i] == ' ';
            }

            return false;
        }

        /// <summary>
        /// Normalizes the span by trimming (in order):
        ///   1) up to three ' ' characters from the start;
        ///   2) all ' ' from the end;
        ///   3) all '#' from the end;
        ///   4) one ' ' from the end;
        ///   5) normalizing inner whitespace.
        /// </summary>
        /// <param name="span">The span to normalize.</param>
        /// <returns>
        /// The normalized span.
        /// </returns>
        [Pure]
        private static ReadOnlySpan<char> Normalize(in ReadOnlySpan<char> span)
            => span.TrimStart().TrimEnd().TrimEnd('#').TrimEnd();

        /// <inheritdoc />
        [Pure]
        public override string ToString() => $"{new string('#', Level)} {Heading}";

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml() => new XElement($"h{Level}", Heading.ToHtml());

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml()
            => new XElement(W + "p",
                new XElement(W + "pPr",
                    new XElement(W + "pStyle",
                        new XAttribute(W + "val", $"Heading{Level}"))),
                new XElement(W + "r", Heading.ToOpenXml()));

        /// <inheritdoc />
        [Pure]
        public bool Equals(MHeading other) => !(other is null) && Heading.Equals(other.Heading) && Level == other.Level;

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is MHeading heading && Equals(heading);

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => unchecked((397 * Heading.GetHashCode()) ^ Level.GetHashCode());

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="MHeading" /> objects are equal.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==([CanBeNull] MHeading left, [CanBeNull] MHeading right)
            => !(left is null) && !(right is null) && left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="MHeading" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=([CanBeNull] MHeading left, [CanBeNull] MHeading right)
            => left is null || right is null || !left.Equals(right);
    }
}