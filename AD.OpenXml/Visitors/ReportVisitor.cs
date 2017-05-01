using AD.IO;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public sealed class ReportVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="result"></param>
        public ReportVisitor([NotNull] DocxFilePath result) : base(result)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        /// <returns></returns>
        protected override OpenXmlVisitor VisitDocument(OpenXmlVisitor subject)
        {
            return new OpenXmlDocumentVisitor(subject);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="footnoteId"></param>
        /// <returns></returns>
        protected override OpenXmlVisitor VisitFootnotes(OpenXmlVisitor subject, int footnoteId)
        {
            return new OpenXmlFootnoteVisitor(subject, footnoteId);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="documentRelationId"></param>
        /// <returns></returns>
        protected override OpenXmlVisitor VisitDocumentHyperlinks(OpenXmlVisitor subject, int documentRelationId)
        {
            return new OpenXmlDocumentHyperlinkVisitor(subject, documentRelationId);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="footnoteRelationId"></param>
        /// <returns></returns>
        protected override OpenXmlVisitor VisitFootnoteHyperlinks(OpenXmlVisitor subject, int footnoteRelationId)
        {
            return new OpenXmlFootnoteHyperlinkVisitor(subject, footnoteRelationId);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        /// <returns></returns>
        protected override OpenXmlVisitor VisitCharts(OpenXmlVisitor subject)
        {
            return new OpenXmlChartVisitor(subject);
        }
    }
}
