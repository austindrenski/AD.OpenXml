using System;
using System.Linq;
using System.Xml.Linq;
using AD.OpenXml.Properties;
using AD.OpenXml.Visitors;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visits
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public sealed class NumberingVisit : IOpenXmlVisit
    {
        [NotNull]
        private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        [NotNull]
        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        [NotNull]
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// 
        /// </summary>
        public IOpenXmlVisitor Result { get; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="subject"></param>
        public NumberingVisit(IOpenXmlVisitor subject)
        {
            (XElement numbering, XElement documentRelations, XElement contentTypes) = 
                Execute(subject.DocumentRelations, subject.ContentTypes, subject.NextDocumentRelationId);

            Result =
                new OpenXmlVisitor(
                    contentTypes,
                    subject.Document,
                    documentRelations,
                    subject.Footnotes,
                    subject.FootnoteRelations,
                    subject.Styles,
                    numbering,
                    subject.Charts);
        }

        [Pure]
        private static (XElement Numbering, XElement DocumentRelations, XElement ContentTypes) 
            Execute(XElement documentRelation, XElement contentTypes, int documentRelationId)
        {
            if (documentRelation is null)
            {
                throw new ArgumentNullException(nameof(documentRelation));
            }
            if (contentTypes is null)
            {
                throw new ArgumentNullException(nameof(contentTypes));
            }

            XElement numbering =
                XElement.Parse(Resources.Numbering);
            
            XElement modifiedDocumentRelations =
                new XElement(
                    documentRelation.Name,
                    documentRelation.Attributes(),
                    documentRelation.Elements().Where(x => x.Attribute("Tartget")?.Value == "numbering.xml"),
                    new XElement(
                        P + "Relationship",
                        new XAttribute("Id", $"rId{documentRelationId}"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                        new XAttribute("Target", "numbering.xml")));

            XElement modifiedContentTypes =
                new XElement(
                    contentTypes.Name,
                    contentTypes.Attributes(),
                    contentTypes.Elements().Where(x => x.Attribute("PartName")?.Value == "/word/numbering.xml"),
                    new XElement(T + "Override",
                        new XAttribute("PartName", "/word/numbering.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")));

            return (numbering, modifiedDocumentRelations, modifiedContentTypes);
        }
    }
}
