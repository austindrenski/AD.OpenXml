using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using AD.IO.Standard;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Enumerates chart file paths in the target document that start like 'word/charts' but not like 'word/charts/_rels'.
    /// </summary>
    [PublicAPI]
    public static class EnumerateChartPathsExtensions
    {
        /// <summary>
        /// Enumerates chart file paths in the target document that start like 'word/charts' but not like 'word/charts/_rels'.
        /// </summary>
        /// <param name="source">The file from which to enumerate entries.</param>
        /// <returns>An <see cref="IEnumerable{T}"/> of chart paths.</returns>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static IEnumerable<string> EnumerateChartPaths([NotNull] this DocxFilePath source)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            IEnumerable<string> charts;
            using (ZipArchive archive = ZipFile.OpenRead(source))
            {
                charts = archive.Entries
                                .Select(x => x.FullName)
                                .Where(x => x.StartsWith("word/charts"))
                                .Where(x => !x.StartsWith("word/charts/_rels"))
                                .ToArray();
            }
            foreach (string item in charts)
            {
                yield return item;
            }
        }
    }
}