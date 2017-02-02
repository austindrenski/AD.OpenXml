using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    [PublicAPI]
    public static class AddAltChunkToDocumentRelationshipsExtensions
    {
        private static readonly XNamespace R = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="toFilePath"></param>
        /// <param name="altChunkFile"></param>
        /// <param name="altChunkName"></param>
        public static void AddAltChunkToDocumentRelationships(this DocxFilePath toFilePath, DocxFilePath altChunkFile, string altChunkName)
        {
            XElement relationships = toFilePath.ReadAsXml("word/_rels/document.xml.rels");

            relationships.Add(
                new XElement(R + "Relationship",
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"),
                    new XAttribute("Target", $"/word/{altChunkFile.Name}{altChunkFile.Extension}"),
                    new XAttribute("Id", altChunkName)));

            relationships.WriteInto(toFilePath, "word/_rels/document.xml.rels");
        }
    }
}
