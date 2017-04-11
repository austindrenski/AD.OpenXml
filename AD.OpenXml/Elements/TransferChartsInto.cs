using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Elements
{
    /// <summary>
    /// 
    /// </summary>
    [PublicAPI]
    public static class TransferChartsExtensions
    {
        private static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        private static readonly XNamespace P = XNamespaces.OpenXmlPackageRelationships;
        
        private static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        private static readonly XNamespace T = XNamespaces.OpenXmlPackageContentTypes;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <param name="fromFilePath"></param>
        /// <param name="toFilePath"></param>
        /// <returns></returns>
        public static XElement TransferCharts(this XElement element, DocxFilePath fromFilePath, DocxFilePath toFilePath)
        {           
            int nextChartId =
                toFilePath.ReadAsXml("word/_rels/document.xml.rels")
                          .Elements()
                          .Attributes("Id")
                          .Select(x => x.Value.Substring(3))
                          .Select(int.Parse)
                          .DefaultIfEmpty(0)
                          .Max();

            IEnumerable<string> fromChartIds =
                fromFilePath.ReadAsXml()
                            .Descendants(C + "chart")
                            .Attributes(R + "id")
                            .Select(x => x.Value)
                            .ToArray();

            foreach (string fromId in fromChartIds)
            {
                string chartId = $"rId{++nextChartId}";
                element = element.ChangeXAttributeValues(C + "chart", R + "id", fromId, chartId);
                TransferChart(fromFilePath, toFilePath, fromId, chartId);
            }

            return element;
        }

        private static void TransferChart(DocxFilePath fromFilePath, DocxFilePath toFilePath, string fromId, string toId)
        {
            string inputName =
                    fromFilePath.ReadAsXml("word/_rels/document.xml.rels")
                                .Elements()
                                .Where(x => x.Attribute("Id")?.Value == fromId)
                                .Select(x => x.Attribute("Target")?.Value.Substring(7))
                                .Single();

            string inputEmbeddingId =
                    fromFilePath.ReadAsXml($"word/charts/{inputName}")
                                .Element(C + "externalData")?
                                .Attribute(R + "id")?
                                .Value;

            string inputEmbeddingName =
                    fromFilePath.ReadAsXml($"word/charts/_rels/{inputName}.rels")
                                .Elements()
                                .Where(x => x.Attribute("Id")?.Value == inputEmbeddingId)
                                .Select(x => x.Attribute("Target")?.Value.Substring(14))
                                .Single();

            int nextChartNumber;
            using (ZipArchive archive = ZipFile.OpenRead(toFilePath))
            {
                nextChartNumber =
                    archive.Entries
                           .Where(x => x.FullName.StartsWith("word/charts/chart"))
                           .Select(x => x.FullName.Replace("word/charts/chart", null).Replace(".xml", null))
                           .Select(int.Parse)
                           .DefaultIfEmpty(0)
                           .Max() + 1;
            }

            int nextEmbeddingNumber;
            bool isXml = false;
            using (ZipArchive archive = ZipFile.OpenRead(toFilePath))
            {
                nextEmbeddingNumber =
                    archive.Entries
                           .Where(x => x.FullName.StartsWith("word/embeddings/"))
                           .Select(
                               x =>
                               {
                                   string name = null;
                                   if (x.FullName.Contains(".xlsx"))
                                   {
                                       name = x.FullName.Replace("word/embeddings/Microsoft_Excel_Worksheet", null).Replace(".xlsx", null);
                                       isXml = true;
                                   }
                                   else if (x.FullName.Contains(".bin"))
                                   {
                                       name = x.FullName.Replace("word/embeddings/oleObject", null).Replace(".bin", null);
                                   }
                                   return name;
                               })
                           .Select(int.Parse)
                           .DefaultIfEmpty(0)
                           .Max() + 1;
            }

            XElement contentTypes = toFilePath.ReadAsXml("[Content_Types].xml");
            contentTypes.Add(
                new XElement(T + "Override", 
                    new XAttribute("PartName", $"/word/charts/chart{nextChartNumber}.xml"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));
            contentTypes.WriteInto(toFilePath, "[Content_Types].xml");

            XElement chartRelation =
                new XElement(P + "Relationships",
                    new XElement(P + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute(
                            "Type",
                            isXml
                                ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
                                : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject"),
                        new XAttribute(
                            "Target",
                            isXml
                                ? $"../embeddings/Microsoft_Excel_Worksheet{nextEmbeddingNumber}.xlsx"
                                : $"../embeddings/oleObject{nextEmbeddingNumber}.bin")));
            
            chartRelation.WriteInto(toFilePath, $"word/charts/_rels/chart{nextChartNumber}.xml.rels");

            XElement documentRelation = toFilePath.ReadAsXml("word/_rels/document.xml.rels");
            documentRelation.Add(
                new XElement(P + "Relationship",
                    new XAttribute("Id", toId),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                    new XAttribute("Target", $"charts/chart{nextChartNumber}.xml")));
            documentRelation.WriteInto(toFilePath, "word/_rels/document.xml.rels");
                           
            fromFilePath.WriteInto(toFilePath, $"word/charts/{inputName}", $"word/charts/chart{nextChartNumber}.xml");

            fromFilePath.WriteInto(
                toFilePath,
                $"word/embeddings/{inputEmbeddingName}",
                isXml
                    ? $"word/embeddings/Microsoft_Excel_Worksheet{nextEmbeddingNumber}.xlsx"
                    : $"word/embeddings/oleObject{nextEmbeddingNumber}.bin");
        }
    }
}