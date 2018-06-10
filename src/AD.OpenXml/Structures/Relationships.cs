using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Structures
{
    // TODO: document ContentTypes
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public class Relationships
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly Relationships Empty = new Relationships();

        /// <summary>
        ///
        /// </summary>
        [NotNull] public static readonly string MimeType = "application/vnd.openxmlformats-package.relationships+xml";

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public IEnumerable<Entry> Entries { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="key"></param>
        public Entry this[string key] => Entries.Single(x => x.Id == key);

        /// <summary>
        ///
        /// </summary>
        /// <param name="entries"></param>
        public Relationships(params IEnumerable<Entry>[] entries)
            => Entries =
                entries.Where(x => x != null)
                       .SelectMany(x => x)
                       .ToArray();

        /// <inheritdoc />
        [Pure]
        public override string ToString() => ToXElement().ToString();

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public XElement ToXElement()
            => new XElement(P + "Relationships",
                Entries.OrderBy(x => x).Select(x => x.ToXElement()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="package"></param>
        /// <param name="partName"></param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] Package package, Uri partName)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            using (Stream stream =
                package.PartExists(partName)
                    ? package.GetPart(partName).GetStream()
                    : package.CreatePart(partName, MimeType).GetStream())
            {
                ToXElement().Save(stream);
            }
        }

        /// <inheritdoc cref="IEquatable{T}"/>
        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public readonly struct Entry : IComparable<Entry>, IEquatable<Entry>
        {
            [NotNull] private static readonly Comparer<int> Comparer = Comparer<int>.Default;

            /// <summary>
            ///
            /// </summary>
            public readonly int NumericId;

            /// <summary>
            ///
            /// </summary>
            [NotNull] public readonly string Id;

            /// <summary>
            ///
            /// </summary>
            [NotNull] public readonly string Target;

            /// <summary>
            ///
            /// </summary>
            [NotNull] public readonly string Type;

            /// <summary>
            ///
            /// </summary>
            [NotNull] public readonly string TargetMode;

            /// <summary>
            ///
            /// </summary>
            /// <param name="type"></param>
            /// <param name="id"></param>
            /// <param name="target"></param>
            /// <param name="targetMode"></param>
            public Entry([NotNull] string id, [NotNull] string target, [NotNull] string type, [CanBeNull] string targetMode = default)
            {
                if (id is null)
                    throw new ArgumentNullException(nameof(id));

                if (target is null)
                    throw new ArgumentNullException(nameof(target));

                if (type is null)
                    throw new ArgumentNullException(nameof(type));

                Id = id;
                NumericId = int.Parse(((ReadOnlySpan<char>) id).Slice(3));
                Target = target;
                Type = type;
                TargetMode = targetMode ?? string.Empty;
            }

            /// <summary>
            ///
            /// </summary>
            /// <param name="entry"></param>
            /// <returns>
            ///
            /// </returns>
            [Pure]
            [NotNull]
            public static explicit operator XElement(Entry entry) => entry.ToXElement();

            /// <summary>
            /// Returns the entry as an <see cref="XElement"/>.
            /// </summary>
            /// <returns>
            /// The entry as an <see cref="XElement"/>.
            /// </returns>
            [Pure]
            [NotNull]
            public XElement ToXElement()
                => new XElement(P + "Relationship",
                    new XAttribute("Id", Id),
                    new XAttribute("Type", Type),
                    new XAttribute("Target", Target),
                    TargetMode.Length != 0 ? new XAttribute("TargetMode", TargetMode) : null);

            /// <inheritdoc />
            [Pure]
            public override string ToString() => ToXElement().ToString();

            /// <inheritdoc />
            [Pure]
            public int CompareTo(Entry other) => Comparer.Compare(NumericId, other.NumericId);

            /// <inheritdoc />
            [Pure]
            public bool Equals(Entry other) => Id.Equals(other.Id) && Type.Equals(other.Type) && Target.Equals(other.Target) && TargetMode.Equals(other.TargetMode);

            /// <inheritdoc />
            [Pure]
            public override bool Equals(object obj) => obj is Entry entry && Equals(entry);

            /// <inheritdoc />
            [Pure]
            public override int GetHashCode()
            {
                unchecked
                {
                    int hashCode = Id.GetHashCode();
                    hashCode = (hashCode * 397) ^ Target.GetHashCode();
                    hashCode = (hashCode * 397) ^ Type.GetHashCode();
                    hashCode = (hashCode * 397) ^ TargetMode.GetHashCode();
                    return hashCode;
                }
            }

            /// <summary>
            /// Returns a value that indicates whether the values of two <see cref="Entry" /> objects are equal.
            /// </summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>
            /// True if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.
            /// </returns>
            [Pure]
            public static bool operator ==(Entry left, Entry right) => left.Equals(right);

            /// <summary>
            /// Returns a value that indicates whether two <see cref="Entry" /> objects have different values.
            /// </summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>
            /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
            /// </returns>
            [Pure]
            public static bool operator !=(Entry left, Entry right) => !left.Equals(right);
        }
    }
}