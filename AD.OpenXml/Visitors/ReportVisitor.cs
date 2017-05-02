using System;
using System.Collections.Generic;
using AD.IO;
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

        /// <summary>
        /// Visit and join the component documents into the <see cref="OpenXmlVisitor"/>.
        /// </summary>
        /// <param name="files">
        /// The files to visit.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        public override OpenXmlVisitor Visit(IEnumerable<DocxFilePath> files)
        {
            if (files is null)
            {
                throw new ArgumentNullException(nameof(files));
            }

            return new ReportVisitor(base.Visit(files));
        }

        /// <summary>
        /// Visit and join the component document into the <see cref="OpenXmlVisitor"/>.
        /// </summary>
        /// <param name="file">
        /// The files to visit.
        /// </param>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        public override OpenXmlVisitor Visit(DocxFilePath file)
        {
            if (file is null)
            {
                throw new ArgumentNullException(nameof(file));
            }

            return new ReportVisitor(base.Visit(file));
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

            return new OpenXmlDocumentVisitor(subject);
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

            return new OpenXmlFootnoteVisitor(subject, footnoteId);
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
        protected override OpenXmlVisitor VisitDocumentHyperlinks(OpenXmlVisitor subject, int documentRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new OpenXmlDocumentHyperlinkVisitor(subject, documentRelationId);
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
        protected override OpenXmlVisitor VisitFootnoteHyperlinks(OpenXmlVisitor subject, int footnoteRelationId)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new OpenXmlFootnoteHyperlinkVisitor(subject, footnoteRelationId);
        }

        /// <summary>
        /// Visit the <see cref="OpenXmlVisitor.Charts"/> and <see cref="OpenXmlVisitor.DocumentRelations"/> of the subject to modify charts in the document.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="OpenXmlVisitor"/> to visit.
        /// </param>
        /// <returns>
        /// A new <see cref="OpenXmlVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        protected override OpenXmlVisitor VisitCharts(OpenXmlVisitor subject)
        {
            if (subject is null)
            {
                throw new ArgumentNullException(nameof(subject));
            }

            return new OpenXmlChartVisitor(subject);
        }
    }
}