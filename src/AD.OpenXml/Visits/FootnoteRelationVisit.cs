using System;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Elements;
using AD.OpenXml.Structures;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public static class FootnoteRelationVisit
    {
        [NotNull] private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">
        /// The file from which content is copied.
        /// </param>
        /// <param name="footnoteRelationId">
        ///
        /// </param>
        /// <returns>
        /// The updated document node of the source file.
        /// </returns>
        public static OpenXmlPackageVisitor VisitFootnotesRels([NotNull] this OpenXmlPackageVisitor subject, int footnoteRelationId)
        {
            if (subject is null)
                throw new ArgumentNullException(nameof(subject));

            Footnotes footnotes = Execute(subject.Footnotes, footnoteRelationId);

            return subject.With(footnotes: footnotes);
        }

        [Pure]
        private static Footnotes Execute(Footnotes footnotes, int footnoteRelationId)
        {
            if (footnotes is null)
                throw new ArgumentNullException(nameof(footnotes));

            XElement f = footnotes.Content.RemoveRsidAttributes();

            XElement modifiedFootnotes =
                new XElement(
                    f.Name,
                    f.Attributes(),
                    f.Elements().Select(UpdateElements));

            HyperlinkInfo[] hyperlinks =
                footnotes.Hyperlinks
                         .Select(x => x.WithOffset(footnoteRelationId))
                         .ToArray();

            return new Footnotes(footnotes.RelationId, modifiedFootnotes, hyperlinks);

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