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
    public struct ImageInformation : IEquatable<ImageInformation>
    {
        /// <summary>
        ///
        /// </summary>
        public static IEqualityComparer<ImageInformation> Comparer = new ImageInformationComparer();

        /// <summary>
        ///
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///
        /// </summary>
        public byte[] Image { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="name"></param>
        /// <param name="image"></param>
        public ImageInformation(string name, byte[] image)
        {
            Name = name;
            Image = image;
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return $"Name: {Name}.";
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            unchecked
            {
                return (397 * (Name?.GetHashCode() ?? 0)) ^ (Image?.GetHashCode() ?? 0);
            }
        }

        /// <inheritdoc />
        public bool Equals(ImageInformation other)
        {
            return
                string.Equals(Name, other.Name) &&
                Equals(Image, other.Image);
        }

        /// <inheritdoc />
        public override bool Equals(object obj)
        {
            return
                !ReferenceEquals(null, obj) &&
                obj is ImageInformation information &&
                Equals(information);
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
        public static bool operator !=(ImageInformation left, ImageInformation right)
        {
            return !left.Equals(right);
        }

        /// <inheritdoc />
        private class ImageInformationComparer : IEqualityComparer<ImageInformation>
        {
            public bool Equals(ImageInformation x, ImageInformation y)
            {
                return x.Equals(y);
            }

            public int GetHashCode(ImageInformation obj)
            {
                return obj.GetHashCode();
            }
        }
    }
}