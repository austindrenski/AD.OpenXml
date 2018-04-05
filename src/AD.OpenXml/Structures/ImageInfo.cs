using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

// BUG: Temporary. Should be fixed in .NET Core 2.1.
// ReSharper disable ImpureMethodCallOnReadonlyValueField

namespace AD.OpenXml.Structures
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public readonly struct ImageInfo : IEquatable<ImageInfo>
    {
        [NotNull] private static readonly Regex RegexTarget = new Regex("media/image(?<id>[0-9]+)\\.(?<extension>png|jpeg|svg)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        ///
        /// </summary>
        private static readonly StringSegment SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly ImageInfo[] Empty = new ImageInfo[0];

        /// <summary>
        ///
        /// </summary>
        public readonly StringSegment Extension;

        /// <summary>
        ///
        /// </summary>
        public readonly StringSegment RelationId;

        /// <summary>
        ///
        /// </summary>
        public readonly ReadOnlyMemory<byte> Image;

        /// <summary>
        ///
        /// </summary>
        public int NumericId => int.Parse(RelationId.Substring(3));

        /// <summary>
        ///
        /// </summary>
        public StringSegment Target => $"media/image{RelationId.Subsegment(3)}.{Extension}";


        /// <summary>
        ///
        /// </summary>
        public StringSegment PartName => $"/word/{Target}";

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
        ///  <param name="rId"></param>
        /// <param name="extension"></param>
        /// <param name="image"></param>
        public ImageInfo(StringSegment rId, StringSegment extension, [NotNull] byte[] image)
        {
            if (!rId.StartsWith("rId", StringComparison.Ordinal))
            {
                throw new ArgumentException($"{nameof(rId)} is not a relationship id.");
            }

            if (image is null)
            {
                throw new ArgumentNullException(nameof(image));
            }

            RelationId = rId;
            Extension = extension;
            Image = image.ToArray();
        }

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="rId"></param>
        /// <param name="extension"></param>
        /// <param name="image"></param>
        public ImageInfo(StringSegment rId, StringSegment extension, ReadOnlyMemory<byte> image)
        {
            if (!rId.StartsWith("rId", StringComparison.Ordinal))
            {
                throw new ArgumentException($"{nameof(rId)} is not a relationship id.");
            }

            RelationId = rId;
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
        public static ImageInfo Create(StringSegment rId, StringSegment target, [NotNull] byte[] image)
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

            string extension = m.Groups["extension"].Value;

            return new ImageInfo(rId, extension, image);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public ImageInfo WithOffset(int offset)
        {
            return new ImageInfo($"rId{NumericId + offset}", Extension, Image);
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
        /// <param name="archive">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            using (Stream stream = archive.CreateEntry(PartName.Subsegment(1).Value).Open())
            {
                for (int i = 0; i < Image.Span.Length; i++)
                {
                    stream.WriteByte(Image.Span[i]);
                }
            }
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
                return (397 * Target.GetHashCode()) ^ Image.GetHashCode();
            }
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is ImageInfo information && Equals(information);
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals(ImageInfo other)
        {
            return Equals(Target, other.Target) && Image.Equals(other.Image);
        }

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Structures.ImageInfo" /> objects are equal.
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
        public static bool operator ==(ImageInfo left, ImageInfo right)
        {
            return left.Equals(right);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.ImageInfo" /> objects have different values.
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
        public static bool operator !=(ImageInfo left, ImageInfo right)
        {
            return !left.Equals(right);
        }
    }
}