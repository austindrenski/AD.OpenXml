using System;
using System.Collections.Generic;
using System.Collections.Immutable;
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
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public IImmutableSet<Entry> Entries { get; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="key">
        ///
        /// </param>
        public Entry this[string key] => Entries.Single(x => x.Id == key);

        /// <summary>
        ///
        /// </summary>
        /// <param name="entries">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public Relationships([NotNull] IEnumerable<Entry> entries)
        {
            if (entries is null)
            {
                throw new ArgumentNullException(nameof(entries));
            }

            Entries = entries.ToImmutableHashSet();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="entries">
        ///
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public Relationships([ItemNotNull] params IEnumerable<Entry>[] entries)
        {
            Entries = entries.SelectMany(x => x).ToImmutableHashSet();
        }

        /// <summary>
        /// Returns the dictionary as an <see cref="XElement"/>.
        /// </summary>
        /// <returns>
        /// The dictionary as an <see cref="XElement"/>.
        /// </returns>
        [Pure]
        [NotNull]
        public XElement ToXElement()
        {
            return new XElement(P + "Relationships", Entries.Select(x => x.ToXElement()));
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return ToXElement().ToString();
        }

        /// <inheritdoc cref="IEquatable{T}"/>
        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public readonly struct Entry : IComparable<Entry>, IEquatable<Entry>
        {
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
            /// <param name="type">
            ///
            /// </param>
            /// <param name="id">
            ///
            /// </param>
            /// <param name="target">
            ///
            /// </param>
            /// <param name="targetMode">
            ///
            /// </param>
            /// <exception cref="ArgumentNullException" />
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
            /// <returns></returns>
            [Pure]
            [NotNull]
            public static explicit operator XElement(Entry entry)
            {
                return entry.ToXElement();
            }

            /// <summary>
            /// Returns the entry as an <see cref="XElement"/>.
            /// </summary>
            /// <returns>
            /// The entry as an <see cref="XElement"/>.
            /// </returns>
            [Pure]
            [NotNull]
            public XElement ToXElement()
            {
                return
                    new XElement(
                        P + "Relationship",
                        new XAttribute("Id", Id),
                        new XAttribute("Type", Type),
                        new XAttribute("Target", Target),
                        TargetMode.HasValue ? new XAttribute("TargetMode", TargetMode) : null);
            }

            /// <inheritdoc />
            [Pure]
            [NotNull]
            public override string ToString()
            {
                return ToXElement().ToString();
            }

            /// <inheritdoc />
            [Pure]
            public int CompareTo(Entry other)
            {
                return StringComparer.Ordinal.Compare(Target.Value, other.Target.Value);
            }

            /// <inheritdoc />
            [Pure]
            public bool Equals(Entry other)
            {
                return Id.Equals(other.Id) && Type.Equals(other.Type) && Target.Equals(other.Target) && TargetMode.Equals(other.TargetMode);
            }

            /// <inheritdoc />
            [Pure]
            public override bool Equals([CanBeNull] object obj)
            {
                return obj is Entry entry && Equals(entry);
            }

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

            /// <summary>Returns a value that indicates whether the values of two <see cref="Entry" /> objects are equal.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.</returns>
            [Pure]
            public static bool operator ==(Entry left, Entry right)
            {
                return left.Equals(right);
            }

            /// <summary>Returns a value that indicates whether two <see cref="Entry" /> objects have different values.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>
            [Pure]
            public static bool operator !=(Entry left, Entry right)
            {
                return !left.Equals(right);
            }
        }
    }
}