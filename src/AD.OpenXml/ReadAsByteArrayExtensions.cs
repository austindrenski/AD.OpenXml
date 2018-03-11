using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using AD.IO.Streams;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    // TODO: add to AD.IO library.
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class ReadAsByteArrayExtensions
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="entryPath"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        [CanBeNull]
        public static byte[] ReadAsByteArray([NotNull] this Stream stream, [NotNull] string entryPath)
        {
            if (stream == null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            if (entryPath == null)
            {
                throw new ArgumentNullException(nameof(entryPath));
            }

            if (stream.CanSeek)
            {
                stream.Seek(default, SeekOrigin.Begin);
            }

            using (ZipArchive zipArchive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                ZipArchiveEntry entry = zipArchive.GetEntry(entryPath);

                if (entry is null)
                {
                    return null;
                }

                using (Stream item = entry.Open())
                {
                    MemoryStream ms = new MemoryStream();
                    item.CopyTo(ms);
                    return ms.ToArray();
                }
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="data"></param>
        /// <param name="toStream"></param>
        /// <param name="entryPath"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        [NotNull]
        public static async Task<MemoryStream> WriteInto(this byte[] data, Stream toStream, string entryPath)
        {
            if (data is null)
            {
                throw new ArgumentNullException(nameof(data));
            }

            if (toStream is null)
            {
                throw new ArgumentNullException(nameof(toStream));
            }

            if (entryPath is null)
            {
                throw new ArgumentNullException(nameof(entryPath));
            }

            MemoryStream memoryStream = await toStream.CopyPure();

            using (ZipArchive zipArchive = new ZipArchive(memoryStream, ZipArchiveMode.Update, true))
            {
                ZipArchiveEntry entry = zipArchive.GetEntry(entryPath);

                entry?.Delete();

                using (Stream stream = zipArchive.CreateEntry(entryPath).Open())
                {
                    stream.Write(data, default, data.Length);
                }
            }

            return memoryStream;
        }
    }
}