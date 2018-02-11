using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AD.IO;
using AD.IO.Streams;
using AD.OpenXml.Structures;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Extensions to enumerate chart part names in the target stream.
    /// </summary>
    [PublicAPI]
    public static class EnumerateChartPartNamesExtensions
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull] private const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

        /// <summary>
        /// Enumerates chart part names in the target stream.
        /// </summary>
        /// <param name="stream">
        /// The stream from which to enumerate entries.
        /// </param>
        /// <returns>
        /// An <see cref="IEnumerable{T}"/> of chart part names.
        /// </returns>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static async Task<IEnumerable<string>> EnumerateChartPartNames([NotNull] this Task<MemoryStream> stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            return await EnumerateChartPartNames(await stream);
        }

        /// <summary>
        /// Enumerates chart part names in the target stream.
        /// </summary>
        /// <param name="stream">
        /// The stream from which to enumerate entries.
        /// </param>
        /// <returns>
        /// An <see cref="IEnumerable{T}"/> of chart part names.
        /// </returns>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static async Task<IEnumerable<string>> EnumerateChartPartNames([NotNull] this MemoryStream stream)
        {
            if (stream is null)
            {
                throw new ArgumentNullException(nameof(stream));
            }

            using (MemoryStream result = await stream.CopyPure())
            {
                return
                    result.ReadAsXml(ContentTypesInfo.Path)
                          .Elements(ContentTypesInfo.Elements.Override)
                          .Where(x => (string) x.Attribute(ContentTypesInfo.Attributes.ContentType) == ChartContentType)
                          .Select(x => (string) x.Attribute(ContentTypesInfo.Attributes.PartName))
                          .Select(x => x.Substring(1));
            }
        }
    }
}