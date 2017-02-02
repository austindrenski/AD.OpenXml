using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Properties;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Documents
{
    [PublicAPI]
    public static class AddHeadersExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlPackageContentTypes;

        private static readonly XNamespace R = XNamespaces.OpenXmlPackageRelationships;

        private static readonly XNamespace S = XNamespaces.OpenXmlOfficeDocumentRelationships;

        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        public static void AddHeaders(this DocxFilePath toFilePath, string title)
        {
            // Modify [Content_Types].xml
            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");
            packageRelation.Descendants(C + "Override")
                           .Where(x => x.Attribute("PartName")?.Value.StartsWith("/word/header") ?? false)
                           .Remove();
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");

            // Modify document.xml.rels and grab the current header id number
            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Descendants(R + "Relationship")
                            .Where(x => x.Attribute("Target")?.Value.Contains("header") ?? false)
                            .Remove();
            int currentHeaderId = documentRelation.Elements().Attributes("Id").Select(x => int.Parse(x.Value.Substring(3))).DefaultIfEmpty(0).Max();
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            // Modify document.xml
            XElement document = toFilePath.ReadAsXml("word/document.xml");
            document.Descendants(W + "sectPr")
                    .Elements(W + "headerReference")
                    .Remove();
            document.WriteInto(toFilePath, "word/document.xml");
            
            // Add headers
            toFilePath.AddEvenPageHeader($"rId{++currentHeaderId}");
            toFilePath.AddOddPageHeader($"rId{++currentHeaderId}", title);
        }

        private static void AddEvenPageHeader(this DocxFilePath toFilePath, string headerId)
        {
            XElement element = XElement.Parse(Resources.header1);
            element.WriteInto(toFilePath, "word/header1.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(R + "Relationship",
                    new XAttribute("Id", headerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute("Target", "header1.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml("word/document.xml");
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(W + "headerReference",
                        new XAttribute(W + "type", "even"),
                        new XAttribute(S + "id", headerId)));
            }
            document.WriteInto(toFilePath, "word/document.xml");

            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");
            packageRelation.Add(
                new XElement(C + "Override",
                    new XAttribute("PartName", "/word/header1.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }

        private static void AddOddPageHeader(this DocxFilePath toFilePath, string headerId, string title)
        {
            XElement element = XElement.Parse(string.Format(Resources.header2, title));
            element.WriteInto(toFilePath, "word/header2.xml");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(R + "Relationship",
                    new XAttribute("Id", headerId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"),
                    new XAttribute("Target", "header2.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");

            XElement document = toFilePath.ReadAsXml("word/document.xml");
            foreach (XElement sectionProperties in document.Descendants(W + "sectPr"))
            {
                sectionProperties.AddFirst(
                    new XElement(W + "headerReference",
                        new XAttribute(W + "type", "default"),
                        new XAttribute(S + "id", headerId)));
            }
            document.WriteInto(toFilePath, "word/document.xml");

            XElement packageRelation = toFilePath.ReadAsXml("[Content_Types].xml");
            packageRelation.Add(
                new XElement(C + "Override",
                    new XAttribute("PartName", "/word/header2.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")));
            packageRelation.WriteInto(toFilePath, "[Content_Types].xml");
        }
    }
}