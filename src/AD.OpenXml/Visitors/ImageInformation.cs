using System;
using System.Collections.Generic;
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
        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public static IEqualityComparer<ImageInformation> Comparer = new ImageInformationComparer();

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string Name { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public byte[] Image { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="name"></param>
        /// <param name="image"></param>
        public ImageInformation([NotNull] string name, [NotNull] byte[] image)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            if (image is null)
            {
                throw new ArgumentNullException(nameof(image));
            }

            Name = name;
            Image = image;
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"Name: {Name}";
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
                return (397 * Name.GetHashCode()) ^ Image.GetHashCode();
            }
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals(ImageInformation other)
        {
            return string.Equals(Name, other.Name) && Equals(Image, other.Image);
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is ImageInformation information && Equals(information);
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

        /// <inheritdoc />
        private class ImageInformationComparer : IEqualityComparer<ImageInformation>
        {
            [Pure]
            public bool Equals(ImageInformation x, ImageInformation y)
            {
                return x.Equals(y);
            }

            [Pure]
            public int GetHashCode(ImageInformation obj)
            {
                return obj.GetHashCode();
            }
        }
    }
}