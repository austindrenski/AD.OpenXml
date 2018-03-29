using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <inheritdoc />
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public sealed class FootnoteRelationVisit : IOpenXmlPackageVisit
    {
        [NotNull] private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <inheritdoc />
        public OpenXmlPackageVisitor Result { get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="footnoteRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public FootnoteRelationVisit(OpenXmlPackageVisitor subject, int footnoteRelationId)
        {
            (var footnoteRelations, var footnotes) =
                Execute(subject.Footnotes.RemoveRsidAttributes(), subject.FootnoteRelations.RemoveRsidAttributes(), footnoteRelationId);

            Result =
                new OpenXmlPackageVisitor(
                    subject.ContentTypes,
                    subject.Document,
                    subject.DocumentRelations,
                    footnotes,
                    footnoteRelations,
                    subject.Styles,
                    subject.Numbering,
                    subject.Theme1,
                    subject.Charts,
                    subject.Images);
        }

        [Pure]
        private static (XElement FootnoteRelations, XElement Footnotes) Execute([NotNull] XElement footnotes, [NotNull] XElement footnoteRelations, int footnoteRelationId)
        {
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }

            if (footnoteRelations is null)
            {
                throw new ArgumentNullException(nameof(footnoteRelations));
            }

//            var footnoteRelationMapping =
//                footnoteRelations.RemoveRsidAttributes()
//                                 .Descendants(P + "Relationship")
//                                 .Select(
//                                     x => new
//                                     {
//                                         Id = x.Attribute("Id"),
//                                         Type = x.Attribute("Type"),
//                                         Target = x.Attribute("Target"),
//                                         TargetMode = x.Attribute("TargetMode")
//                                     })
//                                 .OrderBy(x => x.Id.Value.ParseInt())
//                                 .Select(
//                                     (x, i) => new
//                                     {
//                                         oldId = x.Id,
//                                         newId = new XAttribute("Id", $"rId{footnoteRelationId + i}"),
//                                         x.Type,
//                                         x.Target,
//                                         x.TargetMode
//                                     })
//                                 .ToArray();
//
//            XElement modifiedFootnotes = footnotes.Clone();
//
//            foreach (var map in footnoteRelationMapping)
//            {
//                modifiedFootnotes =
//                    modifiedFootnotes.ChangeXAttributeValues(R + "id", (string) map.oldId, (string) map.newId);
//            }
//
//            XElement modifiedFootnoteRelations =
//                new XElement(
//                    footnoteRelations.Name,
//                    footnoteRelationMapping.Select(
//                        x =>
//                            new XElement(
//                                P + "Relationship",
//                                x.newId,
//                                x.Type,
//                                x.Target,
//                                x.TargetMode)));

            string[] lookup =
                footnoteRelations.Elements(P + "Relationship")
                                 .Select(x => int.Parse(x.Attribute("Id").Value.Substring(3)))
                                 .OrderBy(x => x)
                                 .Select((x, i) => $"rId{i + footnoteRelationId}")
                                 .ToArray();

            XElement modifiedFootnoteRelations =
                new XElement(
                    footnoteRelations.Name,
                    footnoteRelations.Attributes(),
                    lookup.Select(x =>)

                    footnoteRelations.Elements()
                                     .Select(UpdateElements));

            XElement modifiedFootnotes =
                new XElement(
                    footnotes.Name,
                    footnotes.Attributes(),
                    footnotes.Elements().Select(UpdateElements));

            return (modifiedFootnoteRelations, modifiedFootnotes);

            XElement UpdateElements(XElement e)
            {
                return
                    new XElement(
                        e.Name,
                        e.Attributes().Select(UpdateAttributes),
                        e.Elements().Select(UpdateElements),
                        e.HasElements ? null : e.Value);
            }

            XAttribute UpdateAttributes(XAttribute a)
            {
                return
                    a.Name == "Id" || a.Name == R + "id"
                        ? new XAttribute(a.Name, lookup[a.Value])
                        : a;
            }
        }
    }
}