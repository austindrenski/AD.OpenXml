using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Extension methods to merge two <see cref="DocxFilePath"/> documents.
    /// </summary>
    [PublicAPI]
    public static class MergeIntoExtensions
    {
        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="source">The file from which content is copied.</param>
        /// <param name="result">The file into which content is copied.</param>
        public static void MergeInto([NotNull] this DocxFilePath source, [NotNull] DocxFilePath result)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }
            if (result is null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            DocxFilePath tempSource = DocxFilePath.Create($"{source}_temp.docx", true);
            tempSource.AddFootnotes();

            XElement sourceDocument =
                source.ReadAsXml()
                      .Process508From()
                      .TransferFootnotes(source, result)
                      .TransferCharts(source, result);

            XElement resultDocument = result.ReadAsXml();

            XElement sourceBody =
                sourceDocument.Elements()
                              .Single(x => x.Name.LocalName.Equals("body"));

            IEnumerable<XElement> sourceContent =
                sourceBody.Elements();

            resultDocument.Elements().First().Add(sourceContent);

            resultDocument.WriteInto(result, "word/document.xml");

            File.Delete(tempSource);
        }
    }
}