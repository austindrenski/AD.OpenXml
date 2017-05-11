using System.Collections.Generic;
using System.Xml.Linq;
using AD.IO.Standard;
using AD.OpenXml.Visitors;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// Represents a visitor or rewriter for OpenXML documents.
    /// </summary>
    [PublicAPI]
    public interface IOpenXmlVisitor
    {
        /// <summary>
        /// word/charts/chart#.xml.
        /// </summary>
        [NotNull]
        IEnumerable<ChartInformation> Charts { get; }

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        [NotNull]
        XElement ContentTypes { get; }

        /// <summary>
        /// word/document.xml
        /// </summary>
        [NotNull]
        XElement Document { get; }

        /// <summary>
        /// word/_rels/document.xml.rels
        /// </summary>
        [NotNull]
        XElement DocumentRelations { get; }

        /// <summary>
        /// word/_rels/footnotes.xml.rels
        /// </summary>
        [NotNull]
        XElement FootnoteRelations { get; }

        /// <summary>
        /// word/footnotes.xml
        /// </summary>
        [NotNull]
        XElement Footnotes { get; }

        /// <summary>
        /// word/styles.xml
        /// </summary>
        [NotNull]
        XElement Styles { get; }

        /// <summary>
        /// word/numbering.xml
        /// </summary>
        [NotNull]
        XElement Numbering { get; }

        /// <summary>
        /// The current document relation number incremented by one.
        /// </summary>
        int NextDocumentRelationId { get; }

        /// <summary>
        /// The current footnote number incremented by one.
        /// </summary>
        int NextFootnoteId { get; }

        /// <summary>
        /// The current footnote relation number incremented by one.
        /// </summary>
        int NextFootnoteRelationId { get; }

        /// <summary>
        /// The current revision number incremented by one.
        /// </summary>
        int NextRevisionId { get; }
        
        /// <summary>
        /// Writes the <see cref="IOpenXmlVisitor"/> to the <see cref="DocxFilePath"/>.
        /// </summary>
        /// <param name="result">
        /// The file to which the <see cref="IOpenXmlVisitor"/> is written.
        /// </param>
        void Save([NotNull] DocxFilePath result);

        /// <summary>
        /// Visit and join the component document into this <see cref="IOpenXmlVisitor"/>.
        /// </summary>
        /// <param name="file">
        /// The files to visit.
        /// </param>
        [Pure]
        [NotNull]
        IOpenXmlVisitor Visit([NotNull] DocxFilePath file);

        /// <summary>
        /// Folds <paramref name="subject"/> into this <see cref="IOpenXmlVisitor"/>.
        /// </summary>
        /// <param name="subject">
        /// The <see cref="IOpenXmlVisitor"/> that is folded into this <see cref="IOpenXmlVisitor"/>.
        /// </param>
        [Pure]
        [NotNull]
        IOpenXmlVisitor Fold([NotNull] IOpenXmlVisitor subject);
        
        /// <summary>
        /// Visit and fold the component documents into this <see cref="IOpenXmlVisitor"/>.
        /// </summary>
        /// <param name="files">
        /// The files to visit.
        /// </param>
        [Pure]
        [NotNull]
        IOpenXmlVisitor VisitAndFold([ItemNotNull][NotNull] IEnumerable<DocxFilePath> files);
    }
}