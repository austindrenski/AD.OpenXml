using System;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

// ReSharper disable ImpureMethodCallOnReadonlyValueField

namespace AD.OpenXml.Structures
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    /// </summary>
    [PublicAPI]
    public readonly struct ChartInfo : IEquatable<ChartInfo>
    {
        [NotNull] private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        ///
        /// </summary>
        public static readonly StringSegment MimeType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

        /// <summary>
        ///
        /// </summary>
        public static readonly StringSegment SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public XElement Chart { get; }

        /// <summary>
        ///
        /// </summary>
        public StringSegment RelationId { get; }

        /// <summary>
        ///
        /// </summary>
        public StringSegment Target => $"charts/chart{RelationId.Subsegment(3)}.xml";

        /// <summary>
        ///
        /// </summary>
        public StringSegment PartName => $"/word/{Target}";

        /// <summary>
        ///
        /// </summary>
        public ContentTypes.Override ContentTypeEntry => new ContentTypes.Override(PartName, MimeType);

        /// <summary>
        ///
        /// </summary>
        public Relationships.Entry RelationshipEntry => new Relationships.Entry(RelationId, Target, SchemaType);

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="rId"></param>
        /// <param name="chart"></param>
        public ChartInfo(StringSegment rId, [NotNull] XElement chart)
        {
            if (!rId.StartsWith("rId", StringComparison.Ordinal))
            {
                throw new ArgumentException($"{nameof(rId)} is not a relationship id.");
            }

            if (chart is null)
            {
                throw new ArgumentNullException(nameof(chart));
            }

            RelationId = rId;
            XElement clone = chart.Clone();
            clone.Descendants(C + "externalData").Remove();
            Chart = clone;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public ChartInfo WithOffset(uint offset)
        {
            return new ChartInfo($"rId{uint.Parse(RelationId.Substring(3)) + offset}", Chart);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public ChartInfo WithRelationId(StringSegment rId)
        {
            return new ChartInfo(rId, Chart);
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"(Id: {RelationId}, Target: {Target})";
        }

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
        {
            unchecked
            {
                return (397 * RelationId.GetHashCode()) ^ Chart.GetHashCode();
            }
        }

        /// <inheritdoc />
        [Pure]
        public bool Equals(ChartInfo other)
        {
            return Equals(RelationId, other.RelationId) && XNode.DeepEquals(Chart, other.Chart);
        }

        /// <inheritdoc />
        [Pure]
        public override bool Equals([CanBeNull] object obj)
        {
            return obj is ChartInfo chart && Equals(chart);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.ChartInfo" /> objects have the same values.
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
        public static bool operator ==(ChartInfo left, ChartInfo right)
        {
            return left.Equals(right);
        }

        /// <summary>
        /// Returns a value that indicates whether two <see cref="T:AD.OpenXml.Structures.ChartInfo" /> objects have different values.
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
        public static bool operator !=(ChartInfo left, ChartInfo right)
        {
            return !left.Equals(right);
        }
    }
}