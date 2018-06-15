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
    public class MListItem : MNode, IEquatable<MListItem>
    {
        /// <summary>
        /// The text of the heading.
        /// </summary>
        [NotNull]
        public MText Item { get; }

        /// <summary>
        /// The level of the header.
        /// </summary>
        public int Level { get; }

        /// <summary>
        /// Constructs an <see cref="MListItem"/>.
        /// </summary>
        /// <param name="text">The raw text of the item.</param>
        public MListItem(in ReadOnlySpan<char> text)
        {
            if (!Accept(text))
                throw new ArgumentException($"Heading must begin with 1-6 '#' characters followed by a ' ' character: '{text.ToString()}'");

            ReadOnlySpan<char> normalized = Normalize(text);
            Level = normalized.IndexOf(' ');
            Item = normalized.Slice(Level + 1).TrimStart();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="span"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public static implicit operator MListItem(in ReadOnlySpan<char> span) => new MListItem(span);

        /// <summary>
        /// Checks if the segment is a well-formed Markdown heading.
        /// </summary>
        /// <param name="span">The span to test.</param>
        /// <returns>
        /// True if the segment is a well-formed Markdown heading; otherwise false.
        /// </returns>
        public static bool Accept(in ReadOnlySpan<char> span)
        {
            ReadOnlySpan<char> trimmed = span.Trim();

            if (trimmed.Length < 1)
                return false;

            if (trimmed[0] == '-')
                return true;

            return char.IsDigit(trimmed[0]) && (trimmed[1] == '.' || trimmed[1] == ')') && trimmed[2] == ' ';
        }

        /// <summary>
        /// Normalizes the segment by trimming (in order):
        ///   1) up to three ' ' characters from the start;
        ///   2) all ' ' from the end;
        ///   3) all '#' from the end;
        ///   4) one ' ' from the end;
        ///   5) normalizing inner whitespace.
        /// </summary>
        /// <param name="span">The span to normalize.</param>
        /// <returns>
        /// The normalized segment.
        /// </returns>
        [Pure]
        private static ReadOnlySpan<char> Normalize(in ReadOnlySpan<char> span) => span;

        /// <inheritdoc />
        [Pure]
        public override string ToString() => Item.ToString();

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml() => new XElement($"h{Level}", Item.ToHtml());

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml()
            => new XElement(W + "p",
                new XElement(W + "pPr",
                    new XElement(W + "pStyle",
                        new XAttribute(W + "val", $"Heading{Level}"))),
                new XElement(W + "r", Item.ToOpenXml()));

        /// <inheritdoc />
        [Pure]
        public bool Equals(MListItem other) => !(other is null) && Item.Equals(other.Item) && Level == other.Level;

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is MListItem node && Equals(node);

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => unchecked((397 * Item.GetHashCode()) ^ Level.GetHashCode());

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Markdown.MListItem" /> objects are equal.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==([CanBeNull] MListItem left, [CanBeNull] MListItem right)
            => !(left is null) && !(right is null) && left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Markdown.MListItem" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=([CanBeNull] MListItem left, [CanBeNull] MListItem right)
            => left is null || right is null || !left.Equals(right);
    }
}