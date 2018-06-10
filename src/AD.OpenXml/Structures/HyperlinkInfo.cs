﻿using System;
using System.IO.Packaging;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    /// </summary>
    [PublicAPI]
    public readonly struct HyperlinkInfo : IEquatable<HyperlinkInfo>
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string RelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public readonly string RelationId;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public Uri Target { get; }

        /// <summary>
        ///
        /// </summary>
        public readonly TargetMode TargetMode;

        /// <summary>
        ///
        /// </summary>
        public readonly int NumericId;

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <param name="target"></param>
        /// <param name="targetMode"></param>
        /// <exception cref="ArgumentNullException" />
        public HyperlinkInfo([NotNull] string rId, [NotNull] string target, TargetMode targetMode)
        {
            if (rId is null)
                throw new ArgumentNullException(nameof(rId));

            if (target is null)
                throw new ArgumentNullException(nameof(target));

            RelationId = rId;
            NumericId = int.Parse(((ReadOnlySpan<char>) rId).Slice(3));
            Target = new Uri(target, UriKind.RelativeOrAbsolute);
            TargetMode = targetMode;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public HyperlinkInfo WithOffset(int offset) => new HyperlinkInfo($"rId{NumericId + offset}", Target.ToString(), TargetMode);

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public HyperlinkInfo WithRelationId([NotNull] string rId) => new HyperlinkInfo(rId, Target.ToString(), TargetMode);

        /// <inheritdoc />
        [Pure]
        public override string ToString() => $"(Id: {RelationId}, Target: {Target}, TargetMode: {TargetMode})";

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = RelationId.GetHashCode();
                hashCode = (hashCode * 397) ^ Target.GetHashCode();
                hashCode = (hashCode * 397) ^ TargetMode.GetHashCode();
                return hashCode;
            }
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is HyperlinkInfo hyperlink && Equals(hyperlink);

        /// <inheritdoc />
        [Pure]
        public bool Equals(HyperlinkInfo other)
            => Equals(RelationId, other.RelationId) && Equals(Target, other.Target) && Equals(TargetMode, other.TargetMode);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="HyperlinkInfo" /> objects have the same values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==(HyperlinkInfo left, HyperlinkInfo right) => left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="HyperlinkInfo" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=(HyperlinkInfo left, HyperlinkInfo right) => !left.Equals(right);
    }
}