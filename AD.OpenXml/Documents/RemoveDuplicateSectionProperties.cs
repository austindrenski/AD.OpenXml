using System;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Extension methods to remove duplicate section properties.
    /// </summary>
    [PublicAPI]
    public static class RemoveDuplicateSectionPropertiesExtensions
    {
        /// <summary>
        /// Removes section properties elements when they are sequential duplicates. 
        /// </summary>
        /// <param name="result"></param>
        public static void RemoveDuplicateSectionProperties([NotNull] this DocxFilePath result)
        {
            if (result is null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            XElement resultDocument = result.ReadAsXml();

            XElement[] sections =
                resultDocument.Descendants()
                              .Where(x => x.Name.LocalName == "sectPr")
                              .ToArray();

            for (int i = 0; i < sections.Length - 1; i++)
            {
                sections[i].Remove();
            }

            resultDocument.WriteInto(result, "word/document.xml");
        }
    }
}