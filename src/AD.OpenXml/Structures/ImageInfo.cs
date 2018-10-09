using System;
using System.IO;
using System.IO.Packaging;
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
        /// Returns a new target URI.
        /// </summary>
        /// <param name="value">The value used to construct the new URI.</param>
        /// <returns>
        /// A new target URI.
        /// </returns>
        [Pure]
        [NotNull]
        public Uri MakeUri([NotNull] string value)
        {
            if (value is null)
                throw new ArgumentNullException(nameof(value));

            ReadOnlySpan<char> span = TargetUri.OriginalString;
            ReadOnlySpan<char> left = span.Slice(default, span.LastIndexOf('/'));
            ReadOnlySpan<char> right = span.Slice(span.LastIndexOf('.'));

            return new Uri($"{left.ToString()}/image{value}{right.ToString()}", UriKind.Relative);
        }

        /// <summary>
        ///
        /// </summary>
        [Pure]
        [NotNull]
        public string ToBase64String() => Convert.ToBase64String(Image.Span);

        /// <inheritdoc />
        [Pure]
        [NotNull]
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

        /// <summary>
        /// Reads the <see cref="ImageInfo"/> from the <paramref name="part" />.
        /// </summary>
        /// <param name="part">The part from which the image is read.</param>
        /// <param name="relationship">The relationship details of the <paramref name="part"/>.</param>
        /// <returns>
        /// The <see cref="ImageInfo"/> of the specified part and relationship.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public static ImageInfo Read([NotNull] PackagePart part, [NotNull] PackageRelationship relationship)
        {
            if (part is null)
                throw new ArgumentNullException(nameof(part));
            if (relationship is null)
                throw new ArgumentNullException(nameof(relationship));

            using (Stream s = part.GetStream())
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    s.CopyTo(ms);
                    ReadOnlySpan<byte> image = ms.ToArray();
                    return new ImageInfo(relationship.Id, relationship.TargetUri, part.ContentType, image);
                }
            }
        }

        /// <summary>
        /// Writes the image data to the <paramref name="part" />.
        /// </summary>
        /// <param name="part">The part to which the element is written.</param>
        /// <exception cref="ArgumentNullException" />
        public void WriteTo([NotNull] PackagePart part)
        {
            if (part is null)
                throw new ArgumentNullException(nameof(part));

            using (Stream stream = part.GetStream(FileMode.Create))
            {
                stream.Write(Image.Span);
            }
        }
    }
}