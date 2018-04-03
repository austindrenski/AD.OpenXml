using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public readonly struct ImageInformation : IEquatable<ImageInformation>
    {
        [NotNull] private static readonly Regex RegexTarget = new Regex("media/image(?<id>[0-9]+)\\.(?<extension>png|jpeg|svg)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        private readonly uint _id;

        private readonly string _extension;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string RelationId => $"rId{_id}";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string Target => $"media/image{_id}.{_extension}";

        /// <summary>
        ///
        /// </summary>
        public readonly ReadOnlyMemory<byte> Image;

        /// <summary>
        ///
        /// </summary>
        public string Base64String => Convert.ToBase64String(Image.Span.ToArray());

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="id"></param>
        /// <param name="extension"></param>
        /// <param name="image"></param>
        public ImageInformation(uint id, [NotNull] string extension, [NotNull] byte[] image)
        {
            if (extension is null)
            {
                throw new ArgumentNullException(nameof(extension));
            }

            if (image is null)
            {
                throw new ArgumentNullException(nameof(image));
            }


            _id = id;
            _extension = extension;
            Image = image.ToArray();
        }

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="id"></param>
        /// <param name="extension"></param>
        /// <param name="image"></param>
        public ImageInformation(uint id, [NotNull] string extension, ReadOnlyMemory<byte> image)
        {
            if (extension is null)
            {
                throw new ArgumentNullException(nameof(extension));
            }

            _id = id;
            _extension = extension;
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
        public static ImageInformation Create([NotNull] string rId, [NotNull] string target, [NotNull] byte[] image)
        {
            if (rId is null)
            {
                throw new ArgumentNullException(nameof(rId));
            }

            if (target is null)
            {
                throw new ArgumentNullException(nameof(target));
            }

            if (!RegexTarget.IsMatch(target))
            {
                throw new ArgumentException(nameof(target));
            }

            if (image is null)
            {
                throw new ArgumentNullException(nameof(image));
            }

            Match m = RegexTarget.Match(target);

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
            return new ImageInformation(_id + offset, _extension, Image);
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"Target: {Target}";
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
            return string.Equals(Target, other.Target) && Image.Equals(other.Image);
        }

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.Visitors.ImageInformation" /> objects are equal.
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
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Visitors.ImageInformation" /> objects have different values.
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