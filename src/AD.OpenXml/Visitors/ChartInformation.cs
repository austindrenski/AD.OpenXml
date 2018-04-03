using System;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    /// </summary>
    [PublicAPI]
    public readonly struct ChartInformation : IEquatable<ChartInformation>
    {
        [NotNull] private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        /// <summary>
        ///
        /// </summary>
        private readonly uint _id;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string RelationId => $"rId{_id}";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string Target => $"charts/chart{_id}.xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public XElement Chart { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public XElement ContentTypeEntry =>
            new XElement(T + "Override",
                new XAttribute("PartName", $"/word/{Target}"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"));

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="id"></param>
        /// <param name="chart"></param>
        private ChartInformation(uint id, [NotNull] XElement chart)
        {
            if (chart is null)
            {
                throw new ArgumentNullException(nameof(chart));
            }

            _id = id;
            XElement clone = chart.Clone();
            clone.Descendants(C + "externalData").Remove();
            Chart = clone;
        }

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="rId"></param>
        /// <param name="chart"></param>
        ///  <returns></returns>
        ///  <exception cref="ArgumentNullException"></exception>
        public static ChartInformation Create([NotNull] string rId, [NotNull] XElement chart)
        {
            if (rId is null)
            {
                throw new ArgumentNullException(nameof(rId));
            }

            if (chart is null)
            {
                throw new ArgumentNullException(nameof(chart));
            }

            uint id = uint.Parse(rId.Substring(3));

            return new ChartInformation(id, chart);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public ChartInformation WithOffset(uint offset)
        {
            return new ChartInformation(_id + offset, Chart);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public ChartInformation WithRelationId([NotNull] string rId)
        {
            if (rId is null)
            {
                throw new ArgumentNullException(nameof(rId));
            }

            return Create(rId, Chart);
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"(Target: {Target}, FileName: {Chart.Attribute("fileName")})";
        }

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            unchecked
            {
                return (397 * Target.GetHashCode()) ^ Chart.GetHashCode();
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
            return string.Equals(Target, other.Target) && XNode.DeepEquals(Chart, other.Chart);
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