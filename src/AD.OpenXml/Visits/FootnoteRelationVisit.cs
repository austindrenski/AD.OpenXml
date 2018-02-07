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
    /// <inheritdoc />
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public sealed class FootnoteRelationVisit : IOpenXmlVisit
    {
        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull]
        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <inheritdoc />
        public IOpenXmlVisitor Result { get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="footnoteRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public FootnoteRelationVisit(IOpenXmlVisitor subject, int footnoteRelationId)
        {
            (var footnotes, var footnoteRelations) = Execute(subject.Footnotes, subject.FootnoteRelations, footnoteRelationId);

            Result =
                new OpenXmlVisitor(
                    subject.ContentTypes,
                    subject.Document,
                    subject.DocumentRelations,
                    footnotes,
                    footnoteRelations,
                    subject.Styles,
                    subject.Numbering,
                    subject.Theme1,
                    subject.Charts);
        }

        [Pure]
        private static (XElement Footnotes, XElement FootnoteRelations) Execute([NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, int footnoteRelationId)
        {
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }

            var footnoteRelationMapping =
                footnoteRelations.RemoveRsidAttributes()
                                 .Descendants(P + "Relationship")
                                 .Select(
                                     x => new
                                     {
                                         Id = x.Attribute("Id"),
                                         Type = x.Attribute("Type"),
                                         Target = x.Attribute("Target"),
                                         TargetMode = x.Attribute("TargetMode")
                                     })
                                 .OrderBy(x => x.Id.Value.ParseInt())
                                 .Select(
                                     (x, i) => new
                                     {
                                         oldId = x.Id,
                                         newId = new XAttribute("Id", $"rId{footnoteRelationId + i}"),
                                         x.Type,
                                         x.Target,
                                         x.TargetMode
                                     })
                                 .ToArray();

            XElement modifiedFootnotes = footnotes.Clone();

            foreach (var map in footnoteRelationMapping)
            {
                modifiedFootnotes =
                    modifiedFootnotes.ChangeXAttributeValues(R + "id", (string) map.oldId, (string) map.newId);
            }


            XElement modifiedFootnoteRelations =
                new XElement(
                    footnoteRelations.Name,
                    footnoteRelationMapping.Select(
                        x =>
                            new XElement(
                                P + "Relationship",
                                x.newId,
                                x.Type,
                                x.Target,
                                x.TargetMode)));

            return (modifiedFootnotes, modifiedFootnoteRelations);
        }
    }
}