using System;
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
        private readonly uint _id;

        /// <summary>
        ///
        /// </summary>
        public static readonly StringSegment SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        /// <summary>
        ///
        /// </summary>
        public StringSegment RelationId => $"rId{_id}";

        /// <summary>
        ///
        /// </summary>
        public uint NumericId => uint.Parse(RelationId.Substring(3));

        /// <summary>
        ///
        /// </summary>
        public StringSegment Target { get; }

        /// <summary>
        ///
        /// </summary>
        public StringSegment TargetMode { get; }

        /// <summary>
        ///
        /// </summary>
        public Relationships.Entry RelationshipEntry => new Relationships.Entry(RelationId, Target, SchemaType, TargetMode);

        /// <summary>
        ///
        /// </summary>
        /// <param name="id"></param>
        /// <param name="target"></param>
        /// <param name="targetMode"></param>
        /// <exception cref="ArgumentNullException"></exception>
        private HyperlinkInfo(uint id, StringSegment target, StringSegment targetMode)
        {
            _id = id;
            Target = target;
            TargetMode = targetMode;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <param name="target"></param>
        /// <param name="targetMode"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        public static HyperlinkInfo Create(StringSegment rId, StringSegment target, StringSegment targetMode)
        {
            uint id = uint.Parse(rId.Substring(3));

            return new HyperlinkInfo(id, target, targetMode);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public HyperlinkInfo WithOffset(uint offset)
        {
            return new HyperlinkInfo(_id + offset, Target, TargetMode);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public HyperlinkInfo WithRelationId([NotNull] string rId)
        {
            if (rId is null)
            {
                throw new ArgumentNullException(nameof(rId));
            }

            return Create(rId, Target, TargetMode);
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"(Id: {RelationId}, Target: {Target}, TargetMode: {TargetMode})";
        }

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            unchecked
            {
                return (397 * _id.GetHashCode()) ^ (397 * Target.GetHashCode()) ^ (397 * TargetMode.GetHashCode());
            }
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is HyperlinkInfo hyperlink && Equals(hyperlink);
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals(HyperlinkInfo other)
        {
            return _id == other._id && Equals(Target, other.Target) && Equals(TargetMode, other.TargetMode);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.HyperlinkInfo" /> objects have the same values.
        /// </summary>
        /// <param name="left">
        /// The first value to compare.
        /// </param>
        /// <param name="right">
        /// The second value to compare.
        /// </param>
        /// <returns>
        /// true if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==(HyperlinkInfo left, HyperlinkInfo right)
        {
            return left.Equals(right);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.HyperlinkInfo" /> objects have different values.
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
        public static bool operator !=(HyperlinkInfo left, HyperlinkInfo right)
        {
            return !left.Equals(right);
        }
    }
}