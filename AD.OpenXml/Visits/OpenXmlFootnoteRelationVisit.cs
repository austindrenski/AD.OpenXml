using System;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.OpenXml.Visitors;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public sealed class OpenXmlFootnoteRelationVisit : IVisit
    {
        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull]
        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// 
        /// </summary>
        public OpenXmlVisitor Result { get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="footnoteRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public OpenXmlFootnoteRelationVisit(OpenXmlVisitor subject, int footnoteRelationId)
        {
            (var footnotes, var footnoteRelations) = Execute(subject.Footnotes, subject.FootnoteRelations, footnoteRelationId);

            Result =
                new OpenXmlVisitor(
                    subject.File,
                    subject.Document,
                    subject.DocumentRelations,
                    subject.ContentTypes,
                    footnotes,
                    footnoteRelations,
                    subject.Charts);
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
                                     .OrderBy(x => x.Value.ParseInt() ?? 0)
                                     .Select(
                                         (x, i) => new
                                         {
                                             oldId = x.Value,
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