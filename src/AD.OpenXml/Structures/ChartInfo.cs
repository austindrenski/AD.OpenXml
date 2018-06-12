using System;
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
        [NotNull] public const string ContentType =
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string RelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public XElement Chart { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull] public readonly string Id;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public Uri TargetUri { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <param name="targetUri"></param>
        /// <param name="chart"></param>
        /// <exception cref="ArgumentNullException" />
        public ChartInfo([NotNull] string rId, [NotNull] Uri targetUri, [NotNull] XElement chart)
        {
            if (rId is null)
                throw new ArgumentNullException(nameof(rId));
            if (targetUri is null)
                throw new ArgumentNullException(nameof(targetUri));
            if (chart is null)
                throw new ArgumentNullException(nameof(chart));

            Id = rId;
            TargetUri = targetUri;
            Chart = chart.Clone().RemoveByAll(C + "externalData");
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString() => $"(Id: {Id}, TargetUri: {TargetUri})";

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => unchecked((397 * Id.GetHashCode()) ^ Chart.GetHashCode());

        /// <inheritdoc />
        [Pure]
        public bool Equals(ChartInfo other) => Equals(Id, other.Id) && XNode.DeepEquals(Chart, other.Chart);

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