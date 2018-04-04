using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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

        [NotNull] private readonly Dictionary<int, Entry> _dictionary;

        /// <summary>
        ///
        /// </summary>
        /// <param name="key"></param>
        public Entry this[int key] => _dictionary[key];

        /// <summary>
        ///
        /// </summary>
        /// <param name="source"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public Relationships([NotNull] IEnumerable<(int key, Entry value)> source)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            _dictionary = new Dictionary<int, Entry>();
            foreach ((int key, Entry value) in source)
            {
                Add(key, value);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="key"></param>
        /// <param name="entry"></param>
        public void Add(int key, Entry entry)
        {
            _dictionary.Add(key, entry);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="oldKey"></param>
        /// <param name="newKey"></param>
        /// <exception cref="KeyNotFoundException" />
        /// <exception cref="ArgumentException" />
        public void Update(int oldKey, int newKey)
        {
            if (!_dictionary.ContainsKey(oldKey))
            {
                throw new KeyNotFoundException(nameof(oldKey));
            }

            if (_dictionary.ContainsKey(newKey))
            {
                throw new ArgumentException("The new key is already in use.");
            }

            Entry entry = this[oldKey];

            _dictionary.Remove(oldKey);

            _dictionary.Add(newKey, entry);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="oldValue"></param>
        /// <param name="newValue"></param>
        /// <exception cref="KeyNotFoundException" />
        /// <exception cref="ArgumentException" />
        public void Update(Entry oldValue, Entry newValue)
        {
            if (!_dictionary.ContainsValue(oldValue))
            {
                throw new ArgumentException("The old value was not found.");
            }

            if (_dictionary.Count(x => x.Value == oldValue) > 1)
            {
                throw new AmbiguousMatchException("The old value matches more than one entry.");
            }

            int key = _dictionary.Single(x => x.Value == oldValue).Key;

            _dictionary.Remove(key);

            _dictionary.Add(key, newValue);
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public Relationships Refresh()
        {
            return new Relationships(_dictionary.OrderBy(x => x.Key).Select((x, i) => (key: i, value: x.Value)));
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
            return new XElement(P + "Relationships", _dictionary.Select(x => x.ToXElement()));
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