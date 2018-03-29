using System;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

// TODO: temporary until StringSegment is a readonly struct in 2.1
// ReSharper disable ImpureMethodCallOnReadonlyValueField

namespace AD.OpenXml.Markdown
{
    /// <summary>
    /// Represents a Markdown node.
    /// </summary>
    /// <remarks>
    /// See: http://spec.commonmark.org
    /// </remarks>
    [PublicAPI]
    public class MNode : IEquatable<MNode>
    {
        protected static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// The raw text of the node.
        /// </summary>
        public readonly StringSegment Text;

        /// <summary>
        /// Constructs an <see cref="MNode"/>.
        /// </summary>
        /// <param name="text">
        /// The raw text of the node.
        /// </param>
        public MNode(in StringSegment text)
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
        public static implicit operator MNode(in StringSegment segment)
        {
            return new MNode(in segment);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="node">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        public static explicit operator string([CanBeNull] MNode node)
        {
            return node?.Text.Value;
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return Text.Value;
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public virtual XNode ToHtml()
        {
            return new XText(Text.Value);
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public virtual XNode ToOpenXml()
        {
            return
                new XElement(W + "t",
                    new XText(Text.Value));
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals(MNode other)
        {
            return Text.Equals(other.Text);
        }

        /// <summary>
        /// Indicates whether the current object is equal to another object of the same type.
        /// </summary>
        /// <param name="other">
        /// An object to compare with this object.</param>
        /// <returns>
        /// true if the current object is equal to the <paramref name="other">other</paramref> parameter; otherwise, false.
        /// </returns>
        [Pure]
        public bool Equals([CanBeNull] in MNode other)
        {
            return !(other is null) && Text.Equals(other.Text);
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is MNode node && Equals(in node);
        }

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            return Text.GetHashCode();
        }

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Markdown.MNode" /> objects are equal.
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
        public static bool operator ==([CanBeNull] in MNode left, [CanBeNull] in MNode right)
        {
            return !(left is null) && !(right is null) && left.Equals(in right);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Markdown.MNode" /> objects have different values.
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
        public static bool operator !=([CanBeNull] in MNode left, [CanBeNull] in MNode right)
        {
            return left is null || right is null || !left.Equals(in right);
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
    }
}