using System;
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
                Execute(
                    subject.Footnotes.RemoveRsidAttributes(),
                    subject.FootnoteRelations.RemoveRsidAttributes(),
                    footnoteRelationId);

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

            XElement modifiedFootnoteRelations =
                new XElement(
                    footnoteRelations.Name,
                    footnoteRelations.Attributes(),
                    footnoteRelations.Elements().Select(UpdateElements));

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
                        ? new XAttribute(a.Name, $"rId{footnoteRelationId + int.Parse(a.Value.Substring(3))}")
                        : a;
            }
        }
    }
}