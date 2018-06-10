using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

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
                throw new ArgumentNullException(nameof(defaults));

            if (overrides is null)
                throw new ArgumentNullException(nameof(overrides));

            Defaults = defaults.ToArray();
            Overrides = overrides.ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        public ContentTypes([ItemNotNull] params IEnumerable<Override>[] overrides)
        {
            if (overrides is null)
                throw new ArgumentNullException(nameof(overrides));

            Defaults = Default.StandardEntries.ToArray();
            Overrides = overrides.SelectMany(x => x).ToArray();
        }

        /// <summary>
        ///
        /// </summary>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public XContainer ToXElement()
            => new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(T + "Types",
                    Defaults.OrderBy(x => x).Select(x => x.ToXElement()),
                    Overrides.OrderBy(x => x).Select(x => x.ToXElement())));

        /// <inheritdoc />
        [Pure]
        public override string ToString() => ToXElement().ToString();

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
                throw new ArgumentNullException(nameof(archive));

            if (!(archive.GetEntry(ContentTypesInfo.Path) is ZipArchiveEntry entry))
                throw new FileNotFoundException(ContentTypesInfo.Path);

            using (StreamWriter writer = new StreamWriter(entry.Open()))
            {
                writer.Write(ToXElement());
            }
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
            public static ReadOnlyMemory<Default> StandardEntries =>
                new[]
                {
                    new Default("docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
                    new Default("xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                    new Default("rels", "application/vnd.openxmlformats-package.relationships+xml"),
                    new Default("xml", "application/xml"),
                    new Default("jpeg", "image/jpeg"),
                    new Default("png", "image/png"),
                    new Default("svg", "image/svg"),
                    new Default("emf", "image/x-emf")
                };

            /// <summary>
            ///
            /// </summary>
            [NotNull] public readonly string Extension;

            /// <summary>
            ///
            /// </summary>
            [NotNull] public readonly string ContentType;

            /// <summary>
            ///
            /// </summary>
            /// <param name="extension"></param>
            /// <param name="contentType"></param>
            /// <exception cref="ArgumentNullException" />
            public Default([NotNull] string extension, [NotNull] string contentType)
            {
                if (extension is null)
                    throw new ArgumentNullException(nameof(extension));

                if (contentType is null)
                    throw new ArgumentNullException(nameof(contentType));

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
                => new XElement(T + "Default",
                    new XAttribute("Extension", Extension),
                    new XAttribute("ContentType", ContentType));

            /// <inheritdoc />
            [Pure]
            public override string ToString() => ToXElement().ToString();

            /// <inheritdoc />
            [Pure]
            public int CompareTo(Default other) => StringComparer.Ordinal.Compare(Extension, other.Extension);

            /// <inheritdoc />
            [Pure]
            public bool Equals(Default other) => Extension.Equals(other.Extension) && ContentType.Equals(other.ContentType);

            /// <inheritdoc />
            [Pure]
            public override bool Equals(object obj) => obj is Default d && Equals(d);

            /// <inheritdoc />
            [Pure]
            public override int GetHashCode() => unchecked((Extension.GetHashCode() * 397) ^ ContentType.GetHashCode());

            /// <summary>Returns a value that indicates whether two <see cref="Default" /> objects have equal values.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.</returns>
            [Pure]
            public static bool operator ==(Default left, Default right) => left.Equals(right);

            /// <summary>Returns a value that indicates whether two <see cref="Default" /> objects have different values.</summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>true if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.</returns>
            [Pure]
            public static bool operator !=(Default left, Default right) => !left.Equals(right);
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
            [NotNull] public readonly string PartName;

            /// <summary>
            ///
            /// </summary>
            [NotNull] public readonly string ContentType;

            /// <summary>
            ///
            /// </summary>
            /// <param name="partName"></param>
            /// <param name="contentType"></param>
            /// <exception cref="ArgumentNullException" />
            public Override([NotNull] string partName, [NotNull] string contentType)
            {
                if (partName is null)
                    throw new ArgumentNullException(nameof(partName));

                if (contentType is null)
                    throw new ArgumentNullException(nameof(contentType));

                PartName = partName;
                ContentType = contentType;
            }

            /// <summary>
            ///
            /// </summary>
            /// <returns>
            ///
            /// </returns>
            [Pure]
            [NotNull]
            public XElement ToXElement()
                => new XElement(T + "Override",
                    new XAttribute("PartName", PartName),
                    new XAttribute("ContentType", ContentType));

            /// <inheritdoc />
            [Pure]
            public override string ToString() => ToXElement().ToString();

            /// <inheritdoc />
            [Pure]
            public int CompareTo(Override other) => StringComparer.Ordinal.Compare(PartName, other.PartName);

            /// <inheritdoc />
            [Pure]
            public bool Equals(Override other) => PartName.Equals(other.PartName) && ContentType.Equals(other.ContentType);

            /// <inheritdoc />
            [Pure]
            public override bool Equals(object obj) => obj is Override over && Equals(over);

            /// <inheritdoc />
            [Pure]
            public override int GetHashCode() => unchecked((397 * PartName.GetHashCode()) ^ ContentType.GetHashCode());

            /// <summary>
            /// Returns a value that indicates whether two <see cref="Override" /> objects have equal values.
            /// </summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>T
            /// True if <paramref name="left" /> and <paramref name="right" /> are equal; otherwise, false.
            /// </returns>
            [Pure]
            public static bool operator ==(Override left, Override right) => left.Equals(right);

            /// <summary>
            /// Returns a value that indicates whether two <see cref="Override" /> objects have different values.
            /// </summary>
            /// <param name="left">The first value to compare.</param>
            /// <param name="right">The second value to compare.</param>
            /// <returns>
            /// True if <paramref name="left" /> and <paramref name="right" /> are not equal; otherwise, false.
            /// </returns>
            [Pure]
            public static bool operator !=(Override left, Override right) => !left.Equals(right);
        }
    }
}