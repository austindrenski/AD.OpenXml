using System;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Markdown
{
    /// <inheritdoc cref="MNode"/>
    /// <inheritdoc cref="IEquatable{T}"/>
    /// <summary>
    /// Represents a Markdown list item node.
    /// </summary>
    /// <remarks>
    /// See: https://spec.commonmark.org/0.28/#lists
    /// </remarks>
    [PublicAPI]
    public abstract class MListItem : MNode, IEquatable<MListItem>
    {
        /// <summary>
        /// The text of the list item.
        /// </summary>
        [NotNull]
        public abstract MText Item { get; }

        /// <summary>
        /// The level of the list item.
        /// </summary>
        public abstract int Level { get; }

        /// <summary>
        /// The list marker character.
        /// </summary>
        public abstract char Marker { get; }

        /// <summary>
        /// Checks if the segment is a well-formed Markdown heading.
        /// </summary>
        /// <param name="span">The span to test.</param>
        /// <returns>
        /// True if the segment is a well-formed Markdown heading; otherwise false.
        /// </returns>
        [Pure]
        public static bool Accept(in ReadOnlySpan<char> span)
        {
            ReadOnlySpan<char> trimmed = span.Trim();

            // List item must be at least '- '.
            if (trimmed.Length < 2)
                return false;

            if (trimmed.StartsWith("- ") ||
                trimmed.StartsWith("* ") ||
                trimmed.StartsWith("+ "))
                return true;

            // Numbered items must be at least '1) ' or '1. '.
            if (trimmed.Length < 3)
                return false;

            // How many digits in the number?
            int index = 0;
            while (char.IsDigit(trimmed[index]))
            {
                if (++index == trimmed.Length)
                    return false;
            }

            ReadOnlySpan<char> afterDigit = trimmed.Slice(index);

            return afterDigit.StartsWith(") ") ||
                   afterDigit.StartsWith(". ");
        }

        /// <inheritdoc />
        [Pure]
        public abstract override string ToString();

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml() => new XElement("li", Item.ToHtml());

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml()
            => new XElement(W + "p",
                new XElement(W + "pPr",
                    new XElement(W + "pStyle",
                        new XAttribute(W + "val", "ListParagraph"))),
                new XElement(W + "r", Item.ToOpenXml()));

        /// <inheritdoc />
        [Pure]
        public bool Equals(MListItem other)
            => !(other is null) && Marker == other.Marker && Level == other.Level && Item.Equals(other.Item);

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is MListItem node && Equals(node);

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => unchecked((397 * Item.GetHashCode()) ^ (397 * Marker.GetHashCode()) ^ Level.GetHashCode());

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