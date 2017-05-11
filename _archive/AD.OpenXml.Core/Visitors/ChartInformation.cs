using System.Collections.Generic;
using System.Xml.Linq;

namespace AD.OpenXml.Core.Visitors
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
        /// /
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// /
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
            return obj is ChartInformation chart && (Name.Equals(chart.Name) && XNode.DeepEquals(Chart, chart.Chart));
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

        /// <summary>
        /// 
        /// </summary>
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