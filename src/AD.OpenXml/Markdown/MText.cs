using System;
using System.Xml.Linq;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;
using AD.ApiExtensions.Primitives;

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
        public readonly StringSegment Text;

        /// <summary>
        /// Constructs an <see cref="MText"/>.
        /// </summary>
        /// <param name="text">The raw text of the node.</param>
        public MText(in StringSegment text) => Text = text.HasValue ? Normalize(in text) : StringSegment.Empty;

        /// <summary>
        ///
        /// </summary>
        /// <param name="segment"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public static implicit operator MText(in StringSegment segment) => new MText(in segment);

        /// <summary>
        /// Normalizes the segment by trimming outer whitespace and reducing inner whitespace.
        /// </summary>
        /// <param name="segment">The segment to normalize.</param>
        /// <returns>
        /// The normalized segment.
        /// </returns>
        private static StringSegment Normalize(in StringSegment segment) => segment.Trim().NormalizeInner(' ');

        /// <inheritdoc />
        [Pure]
        public override string ToString() => Text.Value;

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml() => new XText(Text.Value);

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml()
            => new XElement(W + "t",
                new XText(Text.Value));

        /// <inheritdoc />
        [Pure]
        public bool Equals(MText other) => !(other is null) && Text.Equals(other.Text);

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is MText node && Equals(node);

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => Text.GetHashCode();

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Markdown.MText" /> objects are equal.
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
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Markdown.MText" /> objects have different values.
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