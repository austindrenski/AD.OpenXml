using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

namespace AD.OpenXml.Structures
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public class ContentTypes
    {
        [NotNull] private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public IEnumerable<Default> Defaults { get; }

        /// <summary>
        ///
        /// </summary>
        [NotNull]
        public IEnumerable<Override> Overrides { get; }

        /// <summary>
        ///
        /// </summary>
        public ContentTypes([NotNull] IEnumerable<Default> defaults, [NotNull] IEnumerable<Override> overrides)
        {
            if (defaults is null)
            {
                throw new ArgumentNullException(nameof(defaults));
            }

            if (overrides is null)
            {
                throw new ArgumentNullException(nameof(overrides));
            }

            Defaults = defaults.ToArray();
            Overrides = overrides.ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="overrides"></param>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public static ContentTypes Create([ItemNotNull] params IEnumerable<Override>[] overrides)
        {
            return new ContentTypes(Default.StandardEntries, overrides.SelectMany(x => x));
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        [Pure]
        [NotNull]
        public XElement ToXElement()
        {
            return
                new XElement(
                    T + "Types",
                    Defaults.OrderBy(x => x).Select(x => x.ToXElement()),
                    Overrides.OrderBy(x => x).Select(x => x.ToXElement()));
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString()
        {
            return ToXElement().ToString();
        }

        /// <inheritdoc cref="IComparable{T}"/>
        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public readonly struct Default : IComparable<Default>, IEquatable<Default>
        {
            /// <summary>
            ///
            /// </summary>
            [NotNull]
            public static Default[] StandardEntries =>
                new Default[]
                {
                    new Default("docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
                    new Default("xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    new Default("rels", "application/vnd.openxmlformats-package.relationships+xml"),
                    new Default("xml", "application/xml"),
                    new Default("jpeg", "image/jpeg"),
                    new Default("png", "image/png"),
                    new Default("svg", "image/svg")
                };

            /// <summary>
            ///
            /// </summary>
            public StringSegment Extension { get; }

            /// <summary>
            ///
            /// </summary>
            public StringSegment ContentType { get; }

            /// <summary>
            ///
            /// </summary>
            /// <param name="extension"></param>
            /// <param name="contentType"></param>
            /// <exception cref="ArgumentNullException" />
            public Default(StringSegment extension, StringSegment contentType)
            {
                if (!extension.HasValue)
                {
                    throw new ArgumentNullException(nameof(extension));
                }

                if (!contentType.HasValue)
                {
                    throw new ArgumentNullException(nameof(contentType));
                }

                Extension = extension;
                ContentType = contentType;
            }

            /// <summary>
            ///
            /// </summary>
            /// <returns></returns>
            [Pure]
            [NotNull]
            public XElement ToXElement()
            {
                return
                    new XElement(
                        T + "Default",
                        new XAttribute("Extension", Extension),
                        new XAttribute("ContentType", ContentType));
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
            public int CompareTo(Default other)
            {
                return StringComparer.Ordinal.Compare(Extension.Value, other.Extension.Value);
            }

            /// <inheritdoc />
            [Pure]
            public bool Equals(Default other)
            {
                return Extension.Equals(other.Extension) && ContentType.Equals(other.ContentType);
            }

            /// <inheritdoc />
            [Pure]
            public override bool Equals([CanBeNull] object obj)
            {
                return obj is Default d && Equals(d);
            }

            /// <inheritdoc />
            [Pure]
            public override int GetHashCode()
            {
                unchecked
                {
                    return (Extension.GetHashCode() * 397) ^ ContentType.GetHashCode();
                }
            }

            /// <summary>Returns a value that indicates whether two <see cref="Default" /> objects have equal values.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.</returns>            [Pure]
            public static bool operator ==(Default left, Default right)
            {
                return left.Equals(right);
            }

            /// <summary>Returns a value that indicates whether two <see cref="Default" /> objects have different values.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>            [Pure]
            public static bool operator !=(Default left, Default right)
            {
                return !left.Equals(right);
            }
        }

        /// <inheritdoc cref="IComparable{T}"/>
        /// <summary>
        ///
        /// </summary>
        [PublicAPI]
        public readonly struct Override : IComparable<Override>, IEquatable<Override>
        {
            /// <summary>
            ///
            /// </summary>
            public StringSegment PartName { get; }

            /// <summary>
            ///
            /// </summary>
            public StringSegment ContentType { get; }

            /// <summary>
            ///
            /// </summary>
            /// <param name="partName"></param>
            /// <param name="contentType"></param>
            /// <exception cref="ArgumentNullException" />
            public Override(StringSegment partName, StringSegment contentType)
            {
                if (!partName.HasValue)
                {
                    throw new ArgumentNullException(nameof(partName));
                }

                if (!contentType.HasValue)
                {
                    throw new ArgumentNullException(nameof(contentType));
                }

                PartName = partName;
                ContentType = contentType;
            }

            /// <summary>
            ///
            /// </summary>
            /// <returns></returns>
            [Pure]
            [NotNull]
            public XElement ToXElement()
            {
                return
                    new XElement(
                        T + "Override",
                        new XAttribute("PartName", PartName),
                        new XAttribute("ContentType", ContentType));
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
            public int CompareTo(Override other)
            {
                return StringComparer.Ordinal.Compare(PartName.Value, other.PartName.Value);
            }

            /// <inheritdoc />
            [Pure]
            public bool Equals(Override other)
            {
                return PartName.Equals(other.PartName) && ContentType.Equals(other.ContentType);
            }

            /// <inheritdoc />
            [Pure]
            public override bool Equals([CanBeNull] object obj)
            {
                return obj is Override over && Equals(over);
            }

            /// <inheritdoc />
            [Pure]
            public override int GetHashCode()
            {
                unchecked
                {
                    return (PartName.GetHashCode() * 397) ^ ContentType.GetHashCode();
                }
            }

            /// <summary>Returns a value that indicates whether two <see cref="Override" /> objects have equal values.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.</returns>
            [Pure]
            public static bool operator ==(Override left, Override right)
            {
                return left.Equals(right);
            }

            /// <summary>Returns a value that indicates whether two <see cref="Override" /> objects have different values.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>
            [Pure]
            public static bool operator !=(Override left, Override right)
            {
                return !left.Equals(right);
            }
        }
    }
}