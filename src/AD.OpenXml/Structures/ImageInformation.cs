using System;
using System.Linq;
using System.Text.RegularExpressions;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

namespace AD.OpenXml.Structures
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public readonly struct ImageInformation : IEquatable<ImageInformation>
    {
        [NotNull] private static readonly Regex RegexTarget = new Regex("media/image(?<id>[0-9]+)\\.(?<extension>png|jpeg|svg)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private readonly uint _id;

        /// <summary>
        ///
        /// </summary>
        public static readonly StringSegment SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        /// <summary>
        ///
        /// </summary>
        public StringSegment Extension { get; }

        /// <summary>
        ///
        /// </summary>
        public StringSegment RelationId => $"rId{_id}";

        /// <summary>
        ///
        /// </summary>
        public StringSegment Target => $"media/image{_id}.{Extension}";

        /// <summary>
        ///
        /// </summary>
        public readonly ReadOnlyMemory<byte> Image;

        /// <summary>
        ///
        /// </summary>
        public string Base64String => Convert.ToBase64String(Image.Span.ToArray());

        /// <summary>
        ///
        /// </summary>
        public Relationships.Entry RelationshipEntry => new Relationships.Entry(RelationId, Target, SchemaType);

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="id"></param>
        /// <param name="extension"></param>
        /// <param name="image"></param>
        public ImageInformation(uint id, StringSegment extension, [NotNull] byte[] image)
        {
            if (image is null)
            {
                throw new ArgumentNullException(nameof(image));
            }


            _id = id;
            Extension = extension;
            Image = image.ToArray();
        }

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="id"></param>
        /// <param name="extension"></param>
        /// <param name="image"></param>
        public ImageInformation(uint id, StringSegment extension, ReadOnlyMemory<byte> image)
        {
            _id = id;
            Extension = extension;
            Image = image;
        }

        ///  <summary>
        ///
        ///  </summary>
        /// <param name="rId"></param>
        /// <param name="target"></param>
        ///  <param name="image"></param>
        ///  <returns></returns>
        ///  <exception cref="ArgumentNullException"></exception>
        public static ImageInformation Create(StringSegment rId, StringSegment target, [NotNull] byte[] image)
        {
            if (!RegexTarget.IsMatch(target.Value))
            {
                throw new ArgumentException(nameof(target));
            }

            if (image is null)
            {
                throw new ArgumentNullException(nameof(image));
            }

            Match m = RegexTarget.Match(target.Value);

            uint id = uint.Parse(rId.Substring(3));
            string extension = m.Groups["extension"].Value;

            return new ImageInformation(id, extension, image);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public ImageInformation WithOffset(uint offset)
        {
            return new ImageInformation(_id + offset, Extension, Image);
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"(Id: {RelationId}, Target: {Target})";
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        public override int GetHashCode()
        {
            unchecked
            {
                // ReSharper disable once ImpureMethodCallOnReadonlyValueField
                return (397 * Target.GetHashCode()) ^ Image.GetHashCode();
            }
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is ImageInformation information && Equals(information);
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals(ImageInformation other)
        {
            // ReSharper disable once ImpureMethodCallOnReadonlyValueField
            return Equals(Target, other.Target) && Image.Equals(other.Image);
        }

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Structures.ImageInformation" /> objects are equal.
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
        public static bool operator ==(ImageInformation left, ImageInformation right)
        {
            return left.Equals(right);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.ImageInformation" /> objects have different values.
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
        public static bool operator !=(ImageInformation left, ImageInformation right)
        {
            return !left.Equals(right);
        }
    }
}