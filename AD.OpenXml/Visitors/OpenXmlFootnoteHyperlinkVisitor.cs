using System;
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
    public class OpenXmlFootnoteHyperlinkVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// Active version of 'word/footnotes.xml'.
        /// </summary>
        public override XElement Footnotes { get; }

        /// <summary>
        /// Active version of 'word/_rels/footnotes.xml.rels'.
        /// </summary>
        public override XElement FootnoteRelations { get; }

        /// <summary>
        /// Returns the last footnote hyperlink identifier currently in use by the container.
        /// </summary>
        public override int FootnoteRelationId { get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="footnoteRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public OpenXmlFootnoteHyperlinkVisitor(OpenXmlVisitor subject, int footnoteRelationId) : base(subject.File)
        {
            (Footnotes, FootnoteRelations, FootnoteRelationId) = Execute(subject.Footnotes, subject.FootnoteRelations, footnoteRelationId);
        }

        [Pure]
        private static (XElement Footnotes, XElement FootnoteRelations, int FootnoteRelationId) Execute([NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, int currentFootnoteRelationId)
        {
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }

            XElement nextFootnoteRelations =
                footnoteRelations.RemoveRsidAttributes() ?? new XElement(P + "Relationships");

            nextFootnoteRelations.Elements()
                                 .Where(x => !x.Attribute("Type")?.Value.Contains("hyperlink") ?? true)
                                 .Remove();

            XElement nextFootnotes =
                footnotes.RemoveRsidAttributes() ?? new XElement(W + "footnotes");

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

            XElement modifiedFootnotes = footnotes.Clone();

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