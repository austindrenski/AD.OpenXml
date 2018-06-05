﻿using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

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
        [NotNull]
        public IImmutableSet<Entry> Entries { get; }

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
                          .ToImmutableHashSet();

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
        /// <param name="archive"></param>
        /// <param name="path"></param>
        /// <exception cref="ArgumentNullException" />
        public void Save([NotNull] ZipArchive archive, string path)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            using (Stream stream = archive.GetEntry(path)?.Open() ?? throw new FileNotFoundException(path))
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
            public int NumericId => int.Parse(Id.Substring(3));

            /// <summary>
            ///
            /// </summary>
            public StringSegment Id { get; }

            /// <summary>
            ///
            /// </summary>
            public StringSegment Target { get; }

            /// <summary>
            ///
            /// </summary>
            public StringSegment Type { get; }

            /// <summary>
            ///
            /// </summary>
            public StringSegment TargetMode { get; }

            /// <summary>
            ///
            /// </summary>
            /// <param name="type"></param>
            /// <param name="id"></param>
            /// <param name="target"></param>
            /// <param name="targetMode"></param>
            public Entry(StringSegment id, StringSegment target, StringSegment type, StringSegment targetMode = default)
            {
                Id = id;
                Target = target;
                Type = type;
                TargetMode = targetMode;
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
                    TargetMode.HasValue ? new XAttribute("TargetMode", TargetMode) : null);

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