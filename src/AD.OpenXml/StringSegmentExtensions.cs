using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

namespace AD.OpenXml
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class StringSegmentExtensions
    {
        /// <summary>
        /// Removes up to the specified count of the specified character from both the start and end of the segment.
        /// </summary>
        /// <param name="segment">
        /// The segment to trim.
        /// </param>
        /// <param name="c">
        /// The character to remove.
        /// </param>
        /// <param name="count">
        /// The maximum number of characters to remove.
        /// </param>
        /// <returns>
        /// The trimmed <see cref="StringSegment" />.
        /// </returns>
        [Pure]
        public static StringSegment Trim(in this StringSegment segment, char c, int count = -1)
        {
            return segment.TrimStart(c, count).TrimEnd(c, count);
        }

        /// <summary>
        /// Removes up to the specified count of the specified character from the start of the segment.
        /// </summary>
        /// <param name="segment">
        /// The segment to trim.
        /// </param>
        /// <param name="c">
        /// The character to remove.
        /// </param>
        /// <param name="count">
        /// The maximum number of characters to remove.
        /// </param>
        /// <returns>
        /// The trimmed <see cref="StringSegment" />.
        /// </returns>
        [Pure]
        public static StringSegment TrimStart(in this StringSegment segment, char c, int count = -1)
        {
            int found = 0;

            for (int i = 0; i < segment.Length; i++)
            {
                if (segment[i] != c)
                {
                    break;
                }

                if (segment[i] == c && count < 0 || found < count)
                {
                    found++;
                }
            }

            return segment.Subsegment(found);
        }

        /// <summary>
        /// Removes up to the specified count of the specified character from the end of the segment.
        /// </summary>
        /// <param name="segment">
        /// The segment to trim.
        /// </param>
        /// <param name="c">
        /// The character to remove.
        /// </param>
        /// <param name="count">
        /// The maximum number of characters to remove.
        /// </param>
        /// <returns>
        /// The trimmed <see cref="StringSegment" />.
        /// </returns>
        [Pure]
        public static StringSegment TrimEnd(in this StringSegment segment, char c, int count = -1)
        {
            int found = 0;

            for (int i = segment.Length - 1; i >= 0; i--)
            {
                if (segment[i] != c)
                {
                    break;
                }

                if (segment[i] == c && count < 0 || found < count)
                {
                    found++;
                }
            }

            return segment.Subsegment(default, segment.Length - found);
        }

        /// <summary>
        /// Reduces multiples of the specified character to one.
        /// </summary>
        /// <param name="segment">
        /// The segment to fix.
        /// </param>
        /// <param name="c">
        /// The character to normalize.
        /// </param>
        /// <returns>
        /// A <see cref="StringSegment"/> representing the corrected string.
        /// </returns>
        [Pure]
        public static StringSegment NormalizeInner(in this StringSegment segment, char c)
        {
            int capacity = segment.Length;

            for (int i = 0; i < segment.Length; i++)
            {
                if (segment[i] == c && i + 1 < segment.Length && segment[i + 1] == c)
                {
                    capacity--;
                }
            }

            InplaceStringBuilder sb = new InplaceStringBuilder(capacity);

            for (int i = 0; i < segment.Length; i++)
            {
                if (segment[i] == c && i + 1 < segment.Length && segment[i + 1] == c)
                {
                    continue;
                }

                sb.Append(segment[i]);
            }

            return sb.ToString();
        }
    }
}