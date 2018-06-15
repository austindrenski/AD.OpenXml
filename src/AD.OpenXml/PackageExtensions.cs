using System;
using System.IO;
using System.IO.Packaging;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    // TODO: document PackageExtensions and relocate
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class PackageExtensions
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="bytes"></param>
        /// <param name="access"></param>
        /// <param name="mode"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public static Package ToPackage(this ReadOnlySpan<byte> bytes, FileAccess access = FileAccess.Read, FileMode mode = FileMode.Open)
        {
            MemoryStream ms = new MemoryStream();
            ms.Write(bytes);

            return Package.Open(ms, mode, access);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="bytes"></param>
        /// <param name="access"></param>
        /// <param name="mode"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [NotNull]
        public static Package ToPackage(this ReadOnlyMemory<byte> bytes, FileAccess access = FileAccess.Read, FileMode mode = FileMode.Open)
            => bytes.Span.ToPackage(access, mode);

        /// <summary>
        ///
        /// </summary>
        /// <param name="package"></param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public static MemoryStream ToStream([NotNull] this Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            MemoryStream ms = new MemoryStream();

            using (Package result = Package.Open(ms, FileMode.Create))
            {
                foreach (PackageRelationship relationship in package.GetRelationships())
                {
                    result.CreateRelationship(relationship.TargetUri, relationship.TargetMode, relationship.RelationshipType, relationship.Id);
                }

                foreach (PackagePart part in package.GetParts())
                {
                    if (part.ContentType == "application/vnd.openxmlformats-package.relationships+xml")
                        continue;

                    PackagePart resultPart =
                        result.PartExists(part.Uri)
                            ? result.GetPart(part.Uri)
                            : result.CreatePart(part.Uri, part.ContentType);

                    foreach (PackageRelationship relationship in part.GetRelationships())
                    {
                        resultPart.CreateRelationship(relationship.TargetUri, relationship.TargetMode, relationship.RelationshipType, relationship.Id);
                    }

                    using (Stream partStream = part.GetStream())
                    {
                        using (Stream resultStream = resultPart.GetStream())
                        {
                            partStream.CopyTo(resultStream);
                        }
                    }
                }
            }

            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }

        ///  <summary>
        ///
        ///  </summary>
        /// <param name="package"></param>
        /// <param name="access"></param>
        /// <param name="mode"></param>
        /// <returns>
        ///
        ///  </returns>
        ///  <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public static Package ToPackage([NotNull] this Package package, FileAccess access = FileAccess.Read, FileMode mode = FileMode.Open)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            return Package.Open(package.ToStream(), mode, access);
        }
    }
}