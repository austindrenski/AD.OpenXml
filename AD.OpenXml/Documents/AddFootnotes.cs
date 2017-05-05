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
        private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="toFilePath"></param>
        public static void AddFootnotes(this DocxFilePath toFilePath)
        {
            // Modify [Content_Types].xml
            XElement packageRelation = 
                toFilePath.ReadAsXml("[Content_Types].xml");

            new XElement(
                    packageRelation.Name,
                    packageRelation.Attributes(),
                    packageRelation.Elements().Where(x => !x.Attribute("PartName")?.Value.StartsWith("/word/footnotes") ?? true),
                    new XElement(
                        T + "Override",
                        new XAttribute("PartName", "/word/footnotes.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")))
                .WriteInto(toFilePath, "[Content_Types].xml");

            // Modify document.xml.rels
            XElement documentRelations = 
                toFilePath.ReadAsXml("word/_rels/document.xml.rels");

            new XElement(
                    documentRelations.Name,
                    documentRelations.Attributes(),
                    documentRelations.Elements(),
                    new XElement(
                        P + "Relationship",
                        new XAttribute("Id", $"rId{documentRelations.Elements().Count() + 1}"),
                        new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"),
                        new XAttribute("Target", "footnotes.xml")))
                .WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement.Parse(Resources.footnotes)
                    .WriteInto(toFilePath, "word/footnotes.xml");
        }
    }
}