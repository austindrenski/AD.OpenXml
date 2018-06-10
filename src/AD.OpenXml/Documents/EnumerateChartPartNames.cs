using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Extensions to enumerate chart parts in the target package.
    /// </summary>
    [PublicAPI]
    public static class EnumerateChartPartNamesExtensions
    {
        /// <summary>
        ///
        /// </summary>
        [NotNull] private const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

        /// <summary>
        /// Enumerates chart part names in the target package.
        /// </summary>
        /// <param name="package">The package from which to enumerate charts.</param>
        /// <returns>An <see cref="IEnumerable{T}"/> of chart part names.</returns>
        [Pure]
        [NotNull]
        [ItemNotNull]
        public static IEnumerable<PackagePart> EnumerateChartPartNames([NotNull] this Package package)
        {
            if (package is null)
                throw new ArgumentNullException(nameof(package));

            return package.GetParts().Where(x => x.ContentType == ChartContentType);
        }
    }
}