using System;
using System.IO;
using AD.OpenXml.Visits;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <inheritdoc />
    /// <summary>
    /// Represents a visitor or rewriter for OpenXML documents.
    /// </summary>
    [PublicAPI]
    public sealed class ReportPackageVisitor : OpenXmlPackageVisitor
    {
        /// <inheritdoc />
        /// <summary>
        /// Initialize a <see cref="ReportPackageVisitor"/> based on a default DOCX <see cref="MemoryStream"/>.
        /// </summary>
        public ReportPackageVisitor()
        {
        }

        /// <inheritdoc />
        /// <summary>
        /// Initialize a <see cref="T:AD.OpenXml.ReportPackageVisitor" /> based on the supplied <see cref="T:AD.IO.Paths.DocxFilePath" />.
        /// </summary>
        /// <param name="result">
        /// The base path used to initialize the new <see cref="T:AD.OpenXml.ReportPackageVisitor" />.
        /// </param>
        public ReportPackageVisitor([NotNull] MemoryStream result) : base(result)
        {
        }

        /// <inheritdoc />
        /// <summary>
        /// Initialize a new <see cref="ReportPackageVisitor"/> from the supplied <see cref="OpenXmlPackageVisitor"/>.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlPackageVisitor"/> used to initialize the new <see cref="ReportPackageVisitor"/>.
        /// </param>
        private ReportPackageVisitor([NotNull] IOpenXmlPackageVisitor subject) : base(subject)
        {
        }

        /// <inheritdoc />
        /// <summary>
        ///
        /// </summary>
        /// <param name="subject"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"/>
        protected override IOpenXmlPackageVisitor Create(IOpenXmlPackageVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new ReportPackageVisitor(subject);
        }

        /// <inheritdoc />
        /// <summary>
        /// Visit the <see cref="IOpenXmlPackageVisitor.Document"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="IOpenXmlPackageVisitor"/> to visit.
        /// </param>
        /// <param name="revisionId">
        /// The current revision number incremented by one.
        /// </param>
        /// <returns>
        /// A new <see cref="IOpenXmlPackageVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override IOpenXmlPackageVisitor VisitDocument(IOpenXmlPackageVisitor subject, int revisionId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return Create(new DocumentVisit(subject, revisionId).Result);
        }

        /// <inheritdoc />
        /// <summary>
        /// Visit the <see cref="IOpenXmlPackageVisitor.Footnotes"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="IOpenXmlPackageVisitor"/> to visit.
        /// </param>
        /// <param name="footnoteId">
        /// The current footnote identifier.
        /// </param>
        /// <param name="revisionId">
        /// The current revision number incremented by one.
        /// </param>
        /// <returns>
        /// A new <see cref="IOpenXmlPackageVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override IOpenXmlPackageVisitor VisitFootnotes(IOpenXmlPackageVisitor subject, int footnoteId, int revisionId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new FootnoteVisit(subject, footnoteId, revisionId).Result;
        }

        /// <inheritdoc />
        /// <summary>
        /// Visit the <see cref="IOpenXmlPackageVisitor.Document"/> and <see cref="IOpenXmlPackageVisitor.DocumentRelations"/> of the subject to modify hyperlinks in the main document.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="IOpenXmlPackageVisitor"/> to visit.
        /// </param>
        /// <param name="documentRelationId">
        /// The current document relationship identifier.
        /// </param>
        /// <returns>
        /// A new <see cref="IOpenXmlPackageVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override IOpenXmlPackageVisitor VisitDocumentRelations(IOpenXmlPackageVisitor subject, int documentRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new DocumentRelationVisit(subject, documentRelationId).Result;
        }

        /// <inheritdoc />
        /// <summary>
        /// Visit the <see cref="IOpenXmlPackageVisitor.Footnotes"/> and <see cref="IOpenXmlPackageVisitor.FootnoteRelations"/> of the subject to modify hyperlinks in the main document.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="IOpenXmlPackageVisitor"/> to visit.
        /// </param>
        /// <param name="footnoteRelationId">
        /// The current footnote relationship identifier.
        /// </param>
        /// <returns>
        /// A new <see cref="IOpenXmlPackageVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override IOpenXmlPackageVisitor VisitFootnoteRelations(IOpenXmlPackageVisitor subject, int footnoteRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new FootnoteRelationVisit(subject, footnoteRelationId).Result;
        }

        /// <inheritdoc />
        /// <summary>
        /// Visit the <see cref="OpenXmlPackageVisitor.Styles"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlPackageVisitor"/> to visit.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlPackageVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override IOpenXmlPackageVisitor VisitStyles(IOpenXmlPackageVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new StyleVisit(subject).Result;
        }

        /// <inheritdoc />
        /// <summary>
        /// Visit the <see cref="OpenXmlPackageVisitor.Numbering"/> of the subject.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlPackageVisitor"/> to visit.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlPackageVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override IOpenXmlPackageVisitor VisitNumbering(IOpenXmlPackageVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new NumberingVisit(subject).Result;
        }
    }
}