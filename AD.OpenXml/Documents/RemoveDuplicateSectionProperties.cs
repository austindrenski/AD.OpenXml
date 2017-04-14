using System;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
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
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

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
                resultDocument.Elements(W + "body")
                              .Elements(W + "sectPr")
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

            //for (int i = 0; i < sections.Length - 1; i++)
            //{
            //    sections[i].Remove();
            //}


            resultDocument.WriteInto(result, "word/document.xml");
        }
    }
}