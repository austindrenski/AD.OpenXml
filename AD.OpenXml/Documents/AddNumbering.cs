using System;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Properties;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    /// <summary>
    /// Extension methods to add the styles part to a Word document.
    /// </summary>
    [PublicAPI]
    public static class AddNumberingExtensions
    {
        [NotNull]
        private static readonly XNamespace C = XNamespaces.OpenXmlPackageContentTypes;

        [NotNull]
        private static readonly XNamespace R = XNamespaces.OpenXmlPackageRelationships;

        /// <summary>
        /// Adds the styles part to a Word document.
        /// </summary>
        /// <param name="toFilePath">The file to which styles are added.</param>
        /// <exception cref="ArgumentNullException"/>
        public static void AddNumbering([NotNull] this DocxFilePath toFilePath)
        {
            if (toFilePath is null)
            {
                throw new ArgumentNullException(nameof(toFilePath));
            }

            XElement numbering = XElement.Parse(Resources.Numbering);
            numbering.WriteInto(toFilePath, "word/numbering.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");

            documentRelation.Descendants(R + "Relationship")
                            .Where(x => x.Attribute("Target")?.Value.Contains("numbering") ?? false)
                            .Remove();

            documentRelation.Add(
                new XElement(R + "Relationship",
                    new XAttribute("Id", $"rId{documentRelation.Elements().Count() + 1}"),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"),
                    new XAttribute("Target", "numbering.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");

            packageRelation.Descendants(C + "Override")
                           .Where(x => x.Attribute("PartName")?.Value == "/word/numbering.xml")
                           .Remove();

            packageRelation.Add(
                new XElement(C + "Override",
                    new XAttribute("PartName", "/word/numbering.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }
    }
}