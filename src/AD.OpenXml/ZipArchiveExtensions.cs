using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    // TODO: document ZipArchiveExtensions and relocate
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class ZipArchiveExtensions
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="archive">
        ///
        /// </param>
        /// <param name="tuples">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        public static ZipArchive With([NotNull] this ZipArchive archive, params (string Path, Func<ZipArchive, XElement> Operation)[] tuples)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            ZipArchive result = archive.ToArchive();

            foreach ((string path, Func<ZipArchive, XElement> operation) in tuples)
            {
                XElement output = operation(result);

                using (Stream stream = result.GetEntry(path)?.Open() ?? result.CreateEntry(path).Open())
                {
                    stream.SetLength(0);
                    output.Save(stream);
                }
            }

            return result;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="archive">
        ///
        /// </param>
        /// <param name="mode">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public static ZipArchive ToArchive([NotNull] this ZipArchive archive, ZipArchiveMode mode = ZipArchiveMode.Update)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            return new ZipArchive(archive.ToStream(), mode);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="archive">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public static MemoryStream ToStream([NotNull] this ZipArchive archive)
        {
            if (archive is null)
            {
                throw new ArgumentNullException(nameof(archive));
            }

            MemoryStream ms = new MemoryStream();

            using (ZipArchive writer = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    using (Stream readStream = entry.Open())
                    {
                        using (Stream writeStream = writer.CreateEntry(entry.FullName).Open())
                        {
                            readStream.CopyTo(writeStream);
                        }
                    }
                }
            }

            ms.Seek(0, SeekOrigin.Begin);

            return ms;
        }
    }
}