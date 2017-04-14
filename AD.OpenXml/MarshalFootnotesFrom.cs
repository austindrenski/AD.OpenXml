using System;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public static class MarshalFootnotesFromExtensions
    {
        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <param name="sourceContent">The document node of the source file containing any modifications made to this point.</param>
        /// <param name="currentFootnoteId">The last footnote number currently in use by the container.</param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        public static (XElement SourceContent, XElement SourceFootnotes, int UpdatedFootnoteId)
            MarshalFootnotesFrom([NotNull] this DocxFilePath file, [NotNull] XElement sourceContent, int currentFootnoteId)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            if (sourceContent is null)
            {
                throw new ArgumentNullException(nameof(sourceContent));
            }

            // TODO: Make the return type of ReadAsXml() a nullable singleton.
            try
            {
                file.ReadAsXml("word/footnotes.xml");
            }
            catch
            {
                return (SourceContent: sourceContent, SourceFootnotes: null, UpdatedFootnoteId: currentFootnoteId);
            }

            XElement sourceFootnotes =
                file.ReadAsXml("word/footnotes.xml")
                    .RemoveRsidAttributes();

            sourceFootnotes.Descendants(W + "p")
                           .Attributes()
                           .Remove();

            sourceFootnotes.Descendants(W + "hyperlink")
                           .Remove();

            var footnoteMapping =
                sourceFootnotes.Elements(W + "footnote")
                               .Attributes(W + "id")
                               .Select(x => x.Value)
                               .Select(int.Parse)
                               .Where(x => x > 0)
                               .OrderByDescending(x => x)
                               .Select(
                                   x => new
                                   {
                                       oldId = $"{x}",
                                       newId = $"{currentFootnoteId + x}",
                                       newNumericId = currentFootnoteId + x
                                   })
                               .ToArray();

            foreach (var map in footnoteMapping)
            {
                sourceContent =
                    sourceContent.ChangeXAttributeValues(W + "footnoteReference", W + "Id", map.oldId, map.newId);

                sourceFootnotes =
                    sourceFootnotes.ChangeXAttributeValues(W + "footnote", W + "id", map.oldId, map.newId);
            }

            int newCurrentId = footnoteMapping.Any() ? footnoteMapping.Max(x => x.newNumericId) : currentFootnoteId;

            return (SourceContent: sourceContent, SourceFootnotes: sourceFootnotes, UpdatedFootnoteId: newCurrentId);
        }

    }
}