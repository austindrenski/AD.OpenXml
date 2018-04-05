using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

// BUG: Temporary. Should be fixed in .NET Core 2.1.
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
        private static readonly StringSegment MimeType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

        /// <summary>
        ///
        /// </summary>
        private static readonly StringSegment SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

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
        public readonly StringSegment RelationId;

        /// <summary>
        ///
        /// </summary>
        public int NumericId => int.Parse(RelationId.Substring(3));

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

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId">
        ///
        /// </param>
        /// <param name="chart">
        ///
        /// </param>
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
            Chart = chart.Clone().RemoveByAll(C + "externalData");
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        public ChartInfo WithOffset(int offset)
        {
            return new ChartInfo($"rId{NumericId + offset}", Chart);
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

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return $"(Id: {RelationId}, PartName: {PartName})";
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="archive">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            using (Stream stream = archive.CreateEntry(PartName.Subsegment(1).Value).Open())
            {
                Chart.Save(stream);
            }
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