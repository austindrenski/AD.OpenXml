using System;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public readonly struct ChartInformation : IEquatable<ChartInformation>
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string Name { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public XElement Chart { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="name"></param>
        /// <param name="chart"></param>
        public ChartInformation([NotNull] string name, [NotNull] XElement chart)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            if (chart is null)
            {
                throw new ArgumentNullException(nameof(chart));
            }

            Name = name;
            Chart = chart;
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"(Name: {Name}, FileName: {Chart.Attribute("fileName")})";
        }

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            unchecked
            {
                return (397 * Name.GetHashCode()) ^ Chart.GetHashCode();
            }
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is ChartInformation chart && Equals(chart);
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals(ChartInformation other)
        {
            return string.Equals(Name, other.Name) && XNode.DeepEquals(Chart, other.Chart);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Visitors.ChartInformation" /> objects have the same values.
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
        public static bool operator ==(ChartInformation left, ChartInformation right)
        {
            return left.Equals(right);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Visitors.ChartInformation" /> objects have different values.
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
        public static bool operator !=(ChartInformation left, ChartInformation right)
        {
            return !left.Equals(right);
        }
    }
}