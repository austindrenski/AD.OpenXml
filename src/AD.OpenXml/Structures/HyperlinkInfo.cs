﻿using System;
using System.Linq;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

// ReSharper disable ImpureMethodCallOnReadonlyValueField
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
        private static readonly StringSegment SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly HyperlinkInfo[] Empty = new HyperlinkInfo[0];

        /// <summary>
        ///
        /// </summary>
        public readonly StringSegment RelationId;

        /// <summary>
        ///
        /// </summary>
        public readonly StringSegment Target;

        /// <summary>
        ///
        /// </summary>
        public readonly StringSegment TargetMode;

        /// <summary>
        ///
        /// </summary>
        public int NumericId => int.Parse(RelationId.Substring(3));

        /// <summary>
        ///
        /// </summary>
        public Relationships.Entry RelationshipEntry => new Relationships.Entry(RelationId, Target, SchemaType, TargetMode);

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <param name="target"></param>
        /// <param name="targetMode"></param>
        /// <exception cref="ArgumentNullException" />
        public HyperlinkInfo(StringSegment rId, StringSegment target, StringSegment targetMode)
        {
            if (!rId.StartsWith("rId", StringComparison.Ordinal))
                throw new ArgumentException($"{nameof(rId)} is not a relationship id.");

            RelationId = rId;
            Target = target;
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
        public HyperlinkInfo WithOffset(int offset) => new HyperlinkInfo($"rId{NumericId + offset}", Target, TargetMode);

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public HyperlinkInfo WithRelationId(StringSegment rId) => new HyperlinkInfo(rId, Target, TargetMode);

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
        public bool Equals(HyperlinkInfo other) => Equals(RelationId, other.RelationId) && Equals(Target, other.Target) && Equals(TargetMode, other.TargetMode);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.HyperlinkInfo" /> objects have the same values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==(HyperlinkInfo left, HyperlinkInfo right) => left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.HyperlinkInfo" /> objects have different values.
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