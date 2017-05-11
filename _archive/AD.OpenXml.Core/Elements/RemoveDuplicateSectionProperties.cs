using System;
using System.Linq;
using System.Xml.Linq;

namespace AD.OpenXml.Core.Elements
{
    /// <summary>
    /// Extension methods to remove duplicate section properties.
    /// </summary>
    [PublicAPI]
    public static class RemoveDuplicateSectionPropertiesExtensions
    {
        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static XElement RemoveDuplicateSectionProperties([NotNull] this XElement document)
        {
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            XElement resultDocument = document.Clone();

            XElement[] sections =
                resultDocument.Descendants(W + "sectPr")
                              .ToArray();

            for (int i = 1; i < sections.Length; i++)
            {
                string previous = sections[i - 1].Element(W + "pgSz")?.Attribute(W + "orient")?.Value;
                string current = sections[i].Element(W + "pgSz")?.Attribute(W + "orient")?.Value;

                if (previous == current)
                {
                    sections[i - 1].Remove();
                }
            }

            return resultDocument;
        }
    }
}