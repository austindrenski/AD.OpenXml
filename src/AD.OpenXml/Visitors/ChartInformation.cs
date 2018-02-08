using System.Collections.Generic;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public struct ChartInformation
    {
        /// <summary>
        ///
        /// </summary>
        public static IEqualityComparer<ChartInformation> Comparer = new ChartInformationComparer();

        /// <summary>
        ///
        /// </summary>
        public string Name { get; }

        /// <summary>
        ///
        /// </summary>
        public XElement Chart { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="name"></param>
        /// <param name="chart"></param>
        public ChartInformation(string name, XElement chart)
        {
            Name = name;
            Chart = chart;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="id">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public static string FormatPath(int id)
        {
            return $"word/charts/chart{id}.xml";
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="id">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public static string FormatPartName(int id)
        {
            return $"/{FormatPath(id)}";
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return $"Name: {Name}. {Chart.Attribute("fileName")}.";
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(object obj)
        {
            return obj is ChartInformation chart && Name.Equals(chart.Name) && XNode.DeepEquals(Chart, chart.Chart);
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
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
            public bool Equals(ChartInformation x, ChartInformation y)
            {
                return x.Equals(y);
            }

            public int GetHashCode(ChartInformation obj)
            {
                return obj.GetHashCode();
            }
        }
    }
}