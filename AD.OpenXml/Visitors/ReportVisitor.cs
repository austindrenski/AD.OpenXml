using System;
using System.Collections.Generic;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Visits;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// Represents a visitor or rewriter for OpenXML documents.
    /// </summary>
    [PublicAPI]
    public sealed class ReportVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// Initialize a <see cref="ReportVisitor"/> based on the supplied <see cref="DocxFilePath"/>.
        /// </summary>
        /// <param name="result">
        /// The base path used to initialize the new <see cref="ReportVisitor"/>.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        public ReportVisitor([NotNull] DocxFilePath result) : base(result) { }

        /// <summary>
        /// Initialize a new <see cref="ReportVisitor"/> from the supplied <see cref="OpenXmlVisitor"/>.
        /// </summary>
        /// <param name="openXmlVisitor">
        /// The <see cref="OpenXmlVisitor"/> used to initialize the new <see cref="ReportVisitor"/>.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        private ReportVisitor([NotNull] OpenXmlVisitor openXmlVisitor) : base(openXmlVisitor) { }
 
        protected override OpenXmlVisitor Create(OpenXmlVisitor subject)
        {
            return new ReportVisitor(subject);
        }

        protected override OpenXmlVisitor Create(DocxFilePath file, XElement document, XElement documentRelations, XElement contentTypes, XElement footnotes, XElement footnoteRelations, IEnumerable<ChartInformation> charts)
        {
            return
                new ReportVisitor(
                    base.Create(
                        file,
                        document,
                        documentRelations,
                        contentTypes,
                        footnotes,
                        footnoteRelations,
                        charts));
        }

        /// <summary>
        /// Visit the <see cref="OpenXmlVisitor.Document"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override OpenXmlVisitor VisitDocument(OpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new OpenXmlDocumentVisit(subject).Result;
        }

        /// <summary>
        /// Visit the <see cref="OpenXmlVisitor.Footnotes"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <param name="footnoteId">
        /// The current footnote identifier.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override OpenXmlVisitor VisitFootnotes(OpenXmlVisitor subject, int footnoteId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new OpenXmlFootnoteVisit(subject, footnoteId).Result;
        }

        /// <summary>
        /// Visit the <see cref="OpenXmlVisitor.Document"/> and <see cref="OpenXmlVisitor.DocumentRelations"/> of the subject to modify hyperlinks in the main document.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <param name="documentRelationId">
        /// The current document relationship identifier.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override OpenXmlVisitor VisitDocumentRelations(OpenXmlVisitor subject, int documentRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new OpenXmlDocumentRelationVisit(subject, documentRelationId).Result;
        }

        /// <summary>
        /// Visit the <see cref="OpenXmlVisitor.Footnotes"/> and <see cref="OpenXmlVisitor.FootnoteRelations"/> of the subject to modify hyperlinks in the main document.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <param name="footnoteRelationId">
        /// The current footnote relationship identifier.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override OpenXmlVisitor VisitFootnoteRelations(OpenXmlVisitor subject, int footnoteRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new OpenXmlFootnoteRelationVisit(subject, footnoteRelationId).Result;
        }
    }
}