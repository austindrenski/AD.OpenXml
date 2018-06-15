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
    public class MParagraph : MNode, IEquatable<MParagraph>
    {
        /// <summary>
        /// The raw text of the node.
        /// </summary>
        public MText Text { get; }

        /// <summary>
        /// Constructs an <see cref="MParagraph"/>.
        /// </summary>
        /// <param name="text">The raw text of the node.</param>
        public MParagraph(in ReadOnlySpan<char> text) => Text = Normalize(text);

        /// <summary>
        /// Normalizes the segment by trimming outer whitespace and reducing inner whitespace.
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
        public override XNode ToHtml()
            => new XElement("p",
                Text.ToHtml());

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml()
            => new XElement(W + "p",
                new XElement(W + "r",
                    Text.ToOpenXml()));

        /// <inheritdoc />
        [Pure]
        public bool Equals(MParagraph other) => !(other is null) && Text.Equals(other.Text);

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is MParagraph node && Equals(node);

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => Text.GetHashCode();

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Markdown.MParagraph" /> objects are equal.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==([CanBeNull] MParagraph left, [CanBeNull] MParagraph right)
            => !(left is null) && !(right is null) && left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Markdown.MParagraph" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=([CanBeNull] MParagraph left, [CanBeNull] MParagraph right)
            => left is null || right is null || !left.Equals(right);
    }
}