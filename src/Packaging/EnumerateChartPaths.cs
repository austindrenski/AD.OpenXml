using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using AD.IO;
using JetBrains.Annotations;

namespace AD.OpenXml.Packaging
{
    [PublicAPI]
    public static class EnumerateChartPathsExtensions
    {
        public static IEnumerable<string> EnumerateChartPaths(this DocxFilePath file)
        {
            IEnumerable<string> charts;
            using (ZipArchive archive = ZipFile.OpenRead(file))
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
