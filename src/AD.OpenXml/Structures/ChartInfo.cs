using System;
using System.IO;
using System.IO.Packaging;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public readonly struct ChartInfo : IEquatable<ChartInfo>
    {
        [NotNull] private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly string MimeType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly string Namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly ChartInfo[] Empty = new ChartInfo[0];

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public XElement Chart { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull] public readonly string RelationId;

        /// <summary>
        ///
        /// </summary>
        public readonly int NumericId;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string Target => $"charts/chart{NumericId}.xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public Uri PartName => new Uri($"/word/{Target}", UriKind.Relative);

        /// <summary>
        ///
        /// </summary>
        public Relationships.Entry RelationshipEntry => new Relationships.Entry(RelationId, Target, Namespace);

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <param name="chart"></param>
        /// <exception cref="ArgumentNullException" />
        public ChartInfo([NotNull] string rId, [NotNull] XElement chart)
        {
            if (rId is null)
                throw new ArgumentNullException(nameof(rId));

            if (chart is null)
                throw new ArgumentNullException(nameof(chart));

            if (!rId.StartsWith("rId", StringComparison.Ordinal))
                throw new ArgumentException($"{nameof(rId)} is not a relationship id.");

            RelationId = rId;
            NumericId = int.Parse(((ReadOnlySpan<char>) rId).Slice(3));
            Chart = chart.Clone().RemoveByAll(C + "externalData");
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public ChartInfo WithOffset(int offset) => new ChartInfo($"rId{NumericId + offset}", Chart);

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public ChartInfo WithRelationId([NotNull] string rId) => new ChartInfo(rId, Chart);

        /// <inheritdoc />
        [Pure]
        public override string ToString() => $"(Id: {RelationId}, PartName: {PartName})";

        /// <summary>
        ///
        /// </summary>
        /// <param name="package">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            using (Stream stream =
                package.PartExists(PartName)
                    ? package.GetPart(PartName).GetStream()
                    : package.CreatePart(PartName, MimeType).GetStream())
            {
                Chart.Save(stream);
            }
        }

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => unchecked((397 * RelationId.GetHashCode()) ^ Chart.GetHashCode());

        /// <inheritdoc />
        [Pure]
        public bool Equals(ChartInfo other) => Equals(RelationId, other.RelationId) && XNode.DeepEquals(Chart, other.Chart);

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is ChartInfo chart && Equals(chart);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="ChartInfo" /> objects have the same values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==(ChartInfo left, ChartInfo right) => left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="ChartInfo" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=(ChartInfo left, ChartInfo right) => !left.Equals(right);
    }
}