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
    public readonly struct DrawingInfo : IEquatable<DrawingInfo>
    {
        [NotNull] static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        [NotNull] static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        // TODO: move to AD.Xml
        [NotNull] static readonly XNamespace CDR = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string ContentType =
            "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public const string RelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes";

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
        [NotNull]
        public XElement Drawing { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="id"></param>
        /// <param name="targetUri"></param>
        /// <param name="drawing"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public DrawingInfo([NotNull] string id, [NotNull] Uri targetUri, [NotNull] XElement drawing)
        {
            if (id is null)
                throw new ArgumentNullException(nameof(id));
            if (targetUri is null)
                throw new ArgumentNullException(nameof(targetUri));
            if (drawing is null)
                throw new ArgumentNullException(nameof(drawing));

            Id = id;
            TargetUri = targetUri;
            Drawing = drawing.Clone();
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString() => $"(Id: {Id}, TargetUri: {TargetUri})";

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode()
            => unchecked((397 * Id.GetHashCode()) ^ (397 * TargetUri.GetHashCode()) ^ Drawing.GetHashCode());

        /// <inheritdoc />
        [Pure]
        public bool Equals(DrawingInfo other) => Equals(Id, other.Id) && XNode.DeepEquals(Drawing, other.Drawing);

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is DrawingInfo drawing && Equals(drawing);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="DrawingInfo" /> objects have the same values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==(DrawingInfo left, DrawingInfo right) => left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="DrawingInfo" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=(DrawingInfo left, DrawingInfo right) => !left.Equals(right);
    }
}