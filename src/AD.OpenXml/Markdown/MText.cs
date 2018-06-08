using System;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Markdown
{
    /// <inheritdoc cref="MNode"/>
    /// <summary>
    /// Represents a Markdown node.
    /// </summary>
    /// <remarks>
    /// See: http://spec.commonmark.org
    /// </remarks>
    [PublicAPI]
    public class MText : MNode, IEquatable<MText>
    {
        /// <summary>
        /// The raw text of the node.
        /// </summary>
        public readonly ReadOnlyMemory<char> Text;

        /// <summary>
        /// Constructs an <see cref="MText"/>.
        /// </summary>
        /// <param name="text">The raw text of the node.</param>
        public MText(in ReadOnlySpan<char> text) => Text = Normalize(in text).ToArray();

        /// <summary>
        ///
        /// </summary>
        /// <param name="span"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public static implicit operator MText(in ReadOnlySpan<char> span) => new MText(in span);

        /// <summary>
        /// Normalizes the span by trimming outer whitespace and reducing inner whitespace.
        /// </summary>
        /// <param name="span">The span to normalize.</param>
        /// <returns>
        /// The normalized segment.
        /// </returns>
        private static ReadOnlySpan<char> Normalize(in ReadOnlySpan<char> span) => span.Trim();

        /// <inheritdoc />
        [Pure]
        public override string ToString() => Text.ToString();

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml() => new XText(Text.ToString());

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml() => new XElement(W + "t", new XText(Text.ToString()));

        /// <inheritdoc />
        [Pure]
        public bool Equals(MText other) => !(other is null) && Text.Span.SequenceEqual(other.Text.Span);

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is MText node && Equals(node);

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => Text.GetHashCode();

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="MText" /> objects are equal.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==([CanBeNull] MText left, [CanBeNull] MText right)
            => !(left is null) && !(right is null) && left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="MText" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=([CanBeNull] MText left, [CanBeNull] MText right)
            => left is null || right is null || !left.Equals(right);
    }
}