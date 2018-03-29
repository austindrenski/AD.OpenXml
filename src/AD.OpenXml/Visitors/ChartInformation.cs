using System;
using System.Collections.Generic;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public readonly struct ChartInformation
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public static IEqualityComparer<ChartInformation> Comparer = new ChartInformationComparer();

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

        /// <summary>
        ///
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is ChartInformation chart && Name.Equals(chart.Name) && XNode.DeepEquals(Chart, chart.Chart);
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
                return 397 * Name.GetHashCode() + 397 * Chart.GetHashCode();
            }
        }

        /// <inheritdoc />
        private class ChartInformationComparer : IEqualityComparer<ChartInformation>
        {
            [Pure]
            public bool Equals(ChartInformation x, ChartInformation y)
            {
                return x.Equals(y);
            }

            [Pure]
            public int GetHashCode(ChartInformation obj)
            {
                return obj.GetHashCode();
            }
        }
    }
}