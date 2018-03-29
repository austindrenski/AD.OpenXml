using System;
using System.Xml.Linq;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

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
        private StringSegment _text;

        /// <summary>
        /// The raw text of the node.
        /// </summary>
        public ref readonly StringSegment Text => ref _text;

        /// <summary>
        /// Constructs an <see cref="MNode"/>.
        /// </summary>
        /// <param name="text">
        /// The raw text of the node.
        /// </param>
        public MNode(in StringSegment text)
        {
            _text = text.HasValue ? FixWhitespace(in text) : StringSegment.Empty;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static implicit operator MNode(in StringSegment text)
        {
            return new MNode(in text);
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
            return new XText(Text.Value);
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals([CanBeNull] MNode other)
        {
            return Text.Equals(other?.Text);
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is MNode node && Equals(node);
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
        public static bool operator ==(in MNode left, in MNode right)
        {
            return Equals(left, right);
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
        public static bool operator !=(MNode left, MNode right)
        {
            return !Equals(left, right);
        }

        /// <summary>
        /// Trims leading and trailing whitespace then reduces multiple whitespaces to one.
        /// </summary>
        /// <param name="segment">
        /// The segment to fix.
        /// </param>
        /// <returns>
        /// A <see cref="StringSegment"/> representing the corrected string.
        /// </returns>
        [Pure]
        private static StringSegment FixWhitespace(in StringSegment segment)
        {
            StringSegment trimmed = segment.Trim();

            int capacity = trimmed.Length;

            for (int i = 0; i < trimmed.Length; i++)
            {
                // safe because trimmed can't end with whitespace.
                if (trimmed[i] == ' ' && trimmed[i + 1] == ' ')
                {
                    capacity--;
                }
            }

            InplaceStringBuilder sb = new InplaceStringBuilder(capacity);

            for (int i = 0; i < trimmed.Length; i++)
            {
                // safe because trimmed can't end with whitespace.
                if (trimmed[i] == ' ' && trimmed[i + 1] == ' ')
                {
                    continue;
                }

                sb.Append(trimmed[i]);
            }

            return sb.ToString();
        }
    }
}