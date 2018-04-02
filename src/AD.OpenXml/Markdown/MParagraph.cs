using System;
using System.Xml.Linq;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

// TODO: temporary until StringSegment is a readonly struct in 2.1
// ReSharper disable ImpureMethodCallOnReadonlyValueField

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
        /// <param name="text">
        /// The raw text of the node.
        /// </param>
        public MParagraph(in StringSegment text)
        {
            Text = text.HasValue ? Normalize(in text) : StringSegment.Empty;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="segment">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public static implicit operator MParagraph(in StringSegment segment)
        {
            return new MParagraph(in segment);
        }

        /// <summary>
        /// Normalizes the segment by trimming outer whitespace and reducing inner whitespace.
        /// </summary>
        /// <param name="segment">
        /// The segment to normalize.
        /// </param>
        /// <returns>
        /// The normalized segment.
        /// </returns>
        private static StringSegment Normalize(in StringSegment segment)
        {
            return segment.Trim().NormalizeInner(' ');
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return Text.ToString();
        }

        /// <inheritdoc />
        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public override XNode ToHtml()
        {
            return Text.ToHtml();
        }

        /// <inheritdoc />
        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public override XNode ToOpenXml()
        {
            return
                new XElement(W + "p",
                    new XElement(W + "r", Text.ToOpenXml()));
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals([CanBeNull] MParagraph other)
        {
            return !(other is null) && Text.Equals(other.Text);
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is MParagraph node && Equals(node);
        }

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            return Text.GetHashCode();
        }

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Markdown.MParagraph" /> objects are equal.
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
        public static bool operator ==([CanBeNull] MParagraph left, [CanBeNull] MParagraph right)
        {
            return !(left is null) && !(right is null) && left.Equals(right);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Markdown.MParagraph" /> objects have different values.
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
        public static bool operator !=([CanBeNull] MParagraph left, [CanBeNull] MParagraph right)
        {
            return left is null || right is null || !left.Equals(right);
        }
    }
}