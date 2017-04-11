using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Properties;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class AddFootnotesExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlPackageContentTypes;

        private static readonly XNamespace R = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="toFilePath"></param>
        public static void AddFootnotes(this DocxFilePath toFilePath)
        {
            // Modify [Content_Types].xml
            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");

            packageRelation.Descendants(C + "Override")
                           .Where(x => x.Attribute("PartName")?.Value.StartsWith("/word/footnotes") ?? false)
                           .Remove();

            packageRelation.Add(
                new XElement(
                    C + "Override",
                    new XAttribute("PartName", "/word/footnotes.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")));

            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");


            // Modify document.xml.rels
            XElement documentRelations = toFilePath.ReadAsXml("word/_rels/document.xml.rels");

            int currentDocumentRelationId =
                documentRelations.Elements()
                                 .Attributes("Id")
                                 .Select(x => x.Value.Substring(3))
                                 .Select(int.Parse)
                                 .DefaultIfEmpty(0)
                                 .Max();

            documentRelations.Add(
                new XElement(
                    R + "Relationship",
                    new XAttribute("Id", $"rId{++currentDocumentRelationId}"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"),
                    new XAttribute("Target", "footnotes.xml")));
            documentRelations.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement footnoteDocument = XElement.Parse(Resources.footnotes);
            footnoteDocument.WriteInto(toFilePath, "word/footnotes.xml");
        }
    }
}