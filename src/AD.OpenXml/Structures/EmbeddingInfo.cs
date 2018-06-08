using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    /// <inheritdoc cref="IEquatable{T}" />
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    [Obsolete("This is for debugging purposes only. Do not use.", true)]
    public readonly struct EmbeddingInfo : IEquatable<EmbeddingInfo>
    {
        [NotNull] private static readonly Regex RegexTarget =
            new Regex("embeddings/oleObject(?<id>[0-9]+)\\.(?<extension>bin)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        ///
        /// </summary>
        [NotNull] private static readonly string SchemaType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly EmbeddingInfo[] Empty = new EmbeddingInfo[0];

        /// <summary>
        ///
        /// </summary>
        [NotNull] public readonly string Extension;

        /// <summary>
        ///
        /// </summary>
        [NotNull] public readonly string RelationId;

        /// <summary>
        ///
        /// </summary>
        public readonly ReadOnlyMemory<byte> Image;

        /// <summary>
        ///
        /// </summary>
        public readonly int NumericId;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string Target => $"embeddings/oleObject{NumericId}.{Extension}";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string PartName => $"/word/{Target}";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public string Base64String => Convert.ToBase64String(Image.Span.ToArray());

        /// <summary>
        ///
        /// </summary>
        public Relationships.Entry RelationshipEntry => new Relationships.Entry(RelationId, Target, SchemaType);

        /// <summary>
        ///
        /// </summary>
        ///  <param name="rId"></param>
        /// <param name="extension"></param>
        /// <param name="image"></param>
        public EmbeddingInfo([NotNull] string rId, [NotNull] string extension, in ReadOnlySpan<byte> image)
        {
            if (!rId.StartsWith("rId", StringComparison.Ordinal))
                throw new ArgumentException($"{nameof(rId)} is not a relationship id.");

            RelationId = rId;
            NumericId = int.Parse(((ReadOnlySpan<char>) rId).Slice(3));
            Extension = extension;
            Image = image.ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="rId"></param>
        /// <param name="target"></param>
        /// <param name="image"></param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        public static EmbeddingInfo Create([NotNull] string rId, [NotNull] string target, in ReadOnlySpan<byte> image)
        {
            if (rId is null)
                throw new ArgumentNullException(nameof(rId));

            if (target is null)
                throw new ArgumentNullException(nameof(target));

            if (!RegexTarget.IsMatch(target))
                throw new ArgumentException(nameof(target));

            Match m = RegexTarget.Match(target);

            string extension = m.Groups["extension"].Value;

            return new EmbeddingInfo(rId, extension, in image);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="offset"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        public EmbeddingInfo WithOffset(int offset) => new EmbeddingInfo($"rId{NumericId + offset}", Extension, Image.Span);

        /// <summary>
        ///
        /// </summary>
        /// <param name="archive"> </param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] ZipArchive archive)
        {
            if (archive is null)
                throw new ArgumentNullException(nameof(archive));

            using (Stream stream = archive.CreateEntry(PartName.Substring(1)).Open())
            {
                stream.Write(Image.Span);
            }
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString() => $"(Id: {RelationId}, Target: {Target})";

        /// <inheritdoc />
        [Pure]
        public override int GetHashCode() => unchecked((397 * Target.GetHashCode()) ^ Image.GetHashCode());

        /// <inheritdoc />
        [Pure]
        public override bool Equals(object obj) => obj is EmbeddingInfo information && Equals(information);

        /// <inheritdoc />
        [Pure]
        public bool Equals(EmbeddingInfo other) => Equals(Target, other.Target) && Image.Span.SequenceEqual(other.Image.Span);

        /// <summary>
        /// Returns a value that indicates whether the values of two <see cref="EmbeddingInfo" /> objects are equal.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator ==(EmbeddingInfo left, EmbeddingInfo right) => left.Equals(right);

        /// <summary>
        /// Returns a value that indicates whether two <see cref="EmbeddingInfo" /> objects have different values.
        /// </summary>
        /// <param name="left">The first value to compare.</param>
        /// <param name="right">The second value to compare.</param>
        /// <returns>
        /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
        /// </returns>
        [Pure]
        public static bool operator !=(EmbeddingInfo left, EmbeddingInfo right) => !left.Equals(right);
    }
}