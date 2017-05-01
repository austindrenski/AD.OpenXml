using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public class OpenXmlDocumentHyperlinkVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// Active version of 'word/document.xml'.
        /// </summary>
        public override XElement Document { get; }
        
        /// <summary>
        /// Active version of 'word/_rels/document.xml.rels'.
        /// </summary>
        public override XElement DocumentRelations { get; }
        
        /// <summary>
        /// Returns the last document relation identifier in use by the container.
        /// </summary>
        public override int DocumentRelationId { get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public OpenXmlDocumentHyperlinkVisitor(OpenXmlVisitor subject, int documentRelationId) : base(subject)
        {
            (Document, DocumentRelations, DocumentRelationId) = Execute(subject.Document, subject.DocumentRelations, documentRelationId);
        }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="documentRelations"></param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        public static (XElement Document, XElement DocumentRelations, int DocumentRelationId) Execute(XElement document, XElement documentRelations, int documentRelationId)
        {
            XElement nextDocumentRelations =
                documentRelations.RemoveRsidAttributes() ?? new XElement(P + "Relationships");

            XElement nextContents =
                document.RemoveRsidAttributes() ?? new XElement(W + "document");

            var documentRelationMapping =
                nextContents.Descendants(W + "hyperlink")
                            .Attributes(R + "id")
                            .Select(x => x.Value.ParseInt() ?? 0)
                            .OrderByDescending(x => x)
                            .Select(
                                x => new
                                {
                                    oldId = $"rId{x}",
                                    newId = $"rId{x + documentRelationId}",
                                    newNumericId = x + documentRelationId
                                })
                            .ToArray();

            XElement modifiedContents = document.Clone();

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
                    : documentRelationId;

            return (modifiedContents, nextDocumentRelations, updatedFootnoteRelationId);
        }
    }
}