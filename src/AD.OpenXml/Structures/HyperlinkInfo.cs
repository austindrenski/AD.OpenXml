using System;
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
        [NotNull] public readonly string Id;

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
        /// <param name="id"></param>
        /// <param name="targetUri"></param>
        /// <param name="targetMode"></param>
        /// <exception cref="ArgumentNullException" />
        public HyperlinkInfo([NotNull] string id, [NotNull] Uri targetUri, TargetMode targetMode)
        {
            if (id is null)
                throw new ArgumentNullException(nameof(id));
            if (targetUri is null)
                throw new ArgumentNullException(nameof(targetUri));

            Id = id;
            Target = targetUri;
            TargetMode = targetMode;
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString() => $"(Id: {Id}, TargetUri: {Target}, TargetMode: {TargetMode})";

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = Id.GetHashCode();
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
            => Equals(Id, other.Id) && Equals(Target, other.Target) && Equals(TargetMode, other.TargetMode);

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