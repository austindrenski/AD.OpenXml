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
    public static class MarshalContentHyperlinksFromExtensions
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
        /// <param name="sourceContents"></param>
        /// <param name="currentDocumentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        public static (XElement HyperlinkModifiedContent, XElement HyperlinkModifiedDocumentRelations, int UpdatedDocumentRelationId)
            MarshalContentHyperlinksFrom([NotNull] this DocxFilePath file, [NotNull] XElement sourceContents, int currentDocumentRelationId)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }
            if (sourceContents is null)
            {
                throw new ArgumentNullException(nameof(sourceContents));
            }

            XElement nextDocumentRelations =
                file.ReadAsXml("word/_rels/document.xml.rels")?
                    .RemoveRsidAttributes()
                    ?? new XElement(P + "Relationships");

            nextDocumentRelations.Elements()
                                 .Where(x => !x.Attribute("Type")?.Value.Contains("hyperlink") ?? true)
                                 .Remove();

            XElement nextContents =
                file.ReadAsXml()?
                    .RemoveRsidAttributes() ?? new XElement(W + "document");

            var documentRelationMapping =
                nextContents.Descendants(W + "hyperlink")
                            .Attributes(R + "id")
                            .Select(x => x.Value.ParseInt() ?? 0)
                            .OrderByDescending(x => x)
                            .Select(
                                x => new
                                {
                                    oldId = $"rId{x}",
                                    newId = $"rId{x + currentDocumentRelationId}",
                                    newNumericId = x + currentDocumentRelationId
                                })
                            .ToArray();

            XElement modifiedContents = sourceContents.Clone();

            foreach (var map in documentRelationMapping)
            {
                modifiedContents =
                    modifiedContents.ChangeXAttributeValues(W + "hyperlink", R + "id", map.oldId, map.newId);

                nextDocumentRelations =
                    nextDocumentRelations.ChangeXAttributeValues(P + "Relationship", "Id", map.oldId, map.newId);
            }

            int updatedFootnoteRelationId =
                documentRelationMapping.Any()
                    ? documentRelationMapping.Max(x => x.newNumericId)
                    : currentDocumentRelationId;

            return (modifiedContents, nextDocumentRelations, updatedFootnoteRelationId);
        }
    }
}