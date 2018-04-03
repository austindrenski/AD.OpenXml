using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull] private static readonly XName RootName = P + "Relationships";

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
            return new XElement(RootName, _dictionary.Select(x => x.ToXElement()));
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return ToXElement().ToString();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        [Pure]
        [NotNull]
        public static implicit operator Relationships([NotNull] XElement node)
        {
            if (node is null)
            {
                throw new ArgumentNullException(nameof(node));
            }

            if (node.Name != RootName)
            {
                throw new ArgumentException($"Root node is not {RootName}.");
            }

            return new Relationships(node.Elements().Select(Entry.Create));
        }

        /// <inheritdoc cref="IEquatable{T}"/>
        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public readonly struct Entry : IEquatable<Entry>
        {
            private static readonly XName EntryName = P + "Relationship";

            /// <summary>
            ///
            /// </summary>
            [NotNull]
            public string Type { get; }

            /// <summary>
            ///
            /// </summary>
            [NotNull]
            public string Target { get; }

            /// <summary>
            ///
            /// </summary>
            /// <param name="type">
            ///
            /// </param>
            /// <param name="target">
            ///
            /// </param>
            /// <exception cref="ArgumentNullException" />
            public Entry([NotNull] string type, [NotNull] string target)
            {
                if (type is null)
                {
                    throw new ArgumentNullException(nameof(type));
                }

                if (target is null)
                {
                    throw new ArgumentNullException(nameof(target));
                }

                Type = type;
                Target = target;
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
                        EntryName,
                        new XAttribute("Type", Type),
                        new XAttribute("Target", Target));
            }

            /// <inheritdoc />
            [Pure]
            [NotNull]
            public override string ToString()
            {
                return ToXElement().ToString();
            }

            /// <summary>
            ///
            /// </summary>
            /// <param name="node"></param>
            /// <returns></returns>
            /// <exception cref="ArgumentNullException"></exception>
            /// <exception cref="ArgumentException"></exception>
            [Pure]
            public static (int Key, Entry Value) Create([NotNull] XElement node)
            {
                if (node is null)
                {
                    throw new ArgumentNullException(nameof(node));
                }

                if (node.Name != EntryName)
                {
                    throw new ArgumentException($"Child node is not {EntryName}.");
                }

                return ((int) node.Attribute("Id"), new Entry((string) node.Attribute("PartName"), (string) node.Attribute("ContentType")));
            }

            /// <inheritdoc />
            public bool Equals(Entry other)
            {
                return string.Equals(Type, other.Type) && string.Equals(Target, other.Target);
            }

            /// <inheritdoc />
            public override bool Equals([CanBeNull] object obj)
            {
                return obj is Entry entry && Equals(entry);
            }

            /// <inheritdoc />
            public override int GetHashCode()
            {
                unchecked
                {
                    return (Type.GetHashCode() * 397) ^ Target.GetHashCode();
                }
            }

            /// <summary>Returns a value that indicates whether the values of two <see cref="T:AD.OpenXml.ContentTypes.Entry" /> objects are equal.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if the <paramref name="left" /> and <paramref name="right" /> parameters have the same value; otherwise, false.</returns>
            public static bool operator ==(Entry left, Entry right)
            {
                return left.Equals(right);
            }

            /// <summary>Returns a value that indicates whether two <see cref="T:AD.OpenXml.ContentTypes.Entry" /> objects have different values.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>
            public static bool operator !=(Entry left, Entry right)
            {
                return !left.Equals(right);
            }
        }
    }
}