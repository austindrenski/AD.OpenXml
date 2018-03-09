using System;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Extensions to support recursive node processing.
    /// </summary>
    [PublicAPI]
    public static class RecurseExtensions
    {
        /// <summary>
        /// Recursively clones the element by applying the element predicate at each level.
        /// </summary>
        /// <param name="element">
        /// The current element.
        /// </param>
        /// <param name="predicate">
        /// The predicate applied to the children of the element.
        /// </param>
        /// <returns>
        /// A replica of the element at the current level.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        public static XElement Recurse([NotNull] this XElement element, [NotNull] Func<XElement, bool> predicate)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            if (predicate is null)
            {
                throw new ArgumentNullException(nameof(predicate));
            }

            return
                new XElement(
                    element.Name,
                    element.Attributes(),
                    element.HasElements ? null : new XText(element.Value),
                    element.Elements()
                           .Where(predicate)
                           .Select(x => x.Recurse(predicate)));
        }
    }
}