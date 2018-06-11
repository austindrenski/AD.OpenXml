using System;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public readonly struct ImageInfo : IEquatable<ImageInfo>
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string RelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public readonly string Id;

        /// <summary>
        ///
        /// </summary>
        public readonly ReadOnlyMemory<byte> Image;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public Uri TargetUri { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string ContentType { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="id"></param>
        /// <param name="targetUri"></param>
        /// <param name="contentType"></param>
        /// <param name="image"></param>
        public ImageInfo([NotNull] string id, [NotNull] Uri targetUri, [NotNull] string contentType, in ReadOnlySpan<byte> image)
        {
            if (id is null)
                throw new ArgumentNullException(nameof(id));
            if (targetUri is null)
                throw new ArgumentNullException(nameof(targetUri));
            if (contentType is null)
                throw new ArgumentNullException(nameof(contentType));

            Id = id;
            TargetUri = targetUri;
            ContentType = contentType;
            Image = image.ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        [Pure]
        [NotNull]
        public string ToBase64String() => Convert.ToBase64String(Image.Span);

        /// <inheritdoc />
        [Pure]
        public override string ToString() => $"(Id: {Id}, TargetUri: {TargetUri})";

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => unchecked((397 * TargetUri.GetHashCode()) ^ Image.GetHashCode());

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is ImageInfo image && Equals(image);

        /// <inheritdoc />
        [Pure]
        public bool Equals(ImageInfo other) => Equals(TargetUri, other.TargetUri) && Image.Span.SequenceEqual(other.Image.Span);

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Structures.ImageInfo" /> objects are equal.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==(ImageInfo left, ImageInfo right) => left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.ImageInfo" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=(ImageInfo left, ImageInfo right) => !left.Equals(right);
    }
}