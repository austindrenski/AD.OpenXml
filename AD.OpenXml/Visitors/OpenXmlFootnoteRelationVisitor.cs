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
    public class OpenXmlFootnoteRelationVisitor : OpenXmlVisitor
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
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="footnoteRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public OpenXmlFootnoteRelationVisitor(OpenXmlVisitor subject, int footnoteRelationId) : base(subject)
        {
            (Footnotes, FootnoteRelations) = Execute(subject.Footnotes, subject.FootnoteRelations, footnoteRelationId);
        }

        [Pure]
        private static (XElement Footnotes, XElement FootnoteRelations) Execute([NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, int footnoteRelationId)
        {
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }

            XElement nextFootnoteRelations =
                footnoteRelations.RemoveRsidAttributes() ?? new XElement(P + "Relationships");

            var footnoteRelationMapping =
                nextFootnoteRelations.Descendants(P + "Relationship")
                                     .Attributes("Id")
                                     .Select(x => x.Value.ParseInt() ?? 0)
                                     .OrderBy(x => x)
                                     .Select(
                                         (x, i) => new
                                         {
                                             oldId = $"rId{x}",
                                             newId = $"rId{footnoteRelationId + i}"
                                         })
                                     .ToArray();

            XElement modifiedFootnotes = footnotes.Clone();

            foreach (var map in footnoteRelationMapping)
            {
                modifiedFootnotes =
                    modifiedFootnotes.ChangeXAttributeValues(R + "id", map.oldId, map.newId);

                nextFootnoteRelations =
                    nextFootnoteRelations.ChangeXAttributeValues("Id", map.oldId, map.newId);
            }
            
            return (modifiedFootnotes, nextFootnoteRelations);
        }
    }
}