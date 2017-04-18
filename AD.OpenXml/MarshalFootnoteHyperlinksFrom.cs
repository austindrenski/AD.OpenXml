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
    public static class MarshalFootnoteHyperlinksFromExtensions
    {
        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of [Content_Types].xml
        /// </summary>
        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of document.xml.
        /// </summary>
        [NotNull]
        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="file">The file from which content is copied.</param>
        /// <param name="sourceFootnotes"></param>
        /// <param name="currentFootnoteRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        public static (XElement HyperlinkModifiedFootnotes, XElement HyperlinkModifiedFootnoteRelations, int UpdatedFootnoteRelationId)
            MarshalFootnoteHyperlinksFrom([NotNull] this DocxFilePath file, [NotNull] XElement sourceFootnotes, int currentFootnoteRelationId)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            if (sourceFootnotes is null)
            {
                throw new ArgumentNullException(nameof(sourceFootnotes));
            }

            XElement nextFootnoteRelations =
                file.ReadAsXml("word/_rels/footnotes.xml.rels")?
                    .RemoveRsidAttributes() ?? new XElement(P + "Relationships");

            nextFootnoteRelations.Elements()
                                 .Where(x => !x.Attribute("Type")?.Value.Contains("hyperlink") ?? true)
                                 .Remove();

            XElement nextFootnotes =
                file.ReadAsXml("word/footnotes.xml")?
                    .RemoveRsidAttributes() ?? new XElement(W + "footnotes");

            var footnoteRelationMapping =
                nextFootnotes.Descendants(W + "hyperlink")
                             .Attributes(R + "id")
                             .Select(x => x.Value.ParseInt() ?? 0)
                             .OrderByDescending(x => x)
                             .Select(
                                 x => new
                                 {
                                     oldId = $"rId{x}",
                                     newId = $"rId{x + currentFootnoteRelationId}",
                                     newNumericId = x + currentFootnoteRelationId
                                 })
                             .ToArray();

            XElement modifiedFootnotes = sourceFootnotes.Clone();

            foreach (var map in footnoteRelationMapping)
            {
                modifiedFootnotes =
                    modifiedFootnotes.ChangeXAttributeValues(W + "hyperlink", R + "id", map.oldId, map.newId);

                nextFootnoteRelations =
                    nextFootnoteRelations.ChangeXAttributeValues(P + "Relationship", "Id", map.oldId, map.newId);
            }

            int updatedFootnoteRelationId =
                footnoteRelationMapping.Any()
                    ? footnoteRelationMapping.Max(x => x.newNumericId)
                    : currentFootnoteRelationId;

            return (modifiedFootnotes, nextFootnoteRelations, updatedFootnoteRelationId);
        }
    }
}