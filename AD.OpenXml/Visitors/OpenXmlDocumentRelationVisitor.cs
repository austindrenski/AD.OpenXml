using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml.Visitors
{
    /// <summary>
    /// Marshals footnotes from the 'footnotes.xml' file of a Word document as idiomatic XML objects.
    /// </summary>
    [PublicAPI]
    public class OpenXmlDocumentRelationVisitor : OpenXmlVisitor
    {
        /// <summary>
        /// Active version of 'word/document.xml'.
        /// </summary>
        public override XElement Document { get; }

        /// <summary>
        /// Active version of 'word/_rels/document.xml.rels'.
        /// </summary>
        public override XElement DocumentRelations { get; }

        /// <summary>
        /// Active version of '[Content_Types].xml'.
        /// </summary>
        public override XElement ContentTypes { get; }

        /// <summary>
        /// Active version of word/charts/chart#.xml.
        /// </summary>
        public override IEnumerable<ChartInformation> Charts{ get; }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="subject">The file from which content is copied.</param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        public OpenXmlDocumentRelationVisitor(OpenXmlVisitor subject, int documentRelationId) : base(subject)
        {
            (Document, DocumentRelations, ContentTypes, Charts) = Execute(subject.Document, subject.DocumentRelations, subject.ContentTypes, subject.Charts, documentRelationId);
        }

        /// <summary>
        /// Marshals footnotes from the source document into the container.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="documentRelations"></param>
        /// <param name="contentTypes"></param>
        /// <param name="charts"></param>
        /// <param name="documentRelationId"></param>
        /// <returns>The updated document node of the source file.</returns>
        [Pure]
        private static (XElement Document, XElement DocumentRelations, XElement ContentTypes, IEnumerable<ChartInformation> Charts) Execute(XElement document, XElement documentRelations, XElement contentTypes, IEnumerable<ChartInformation> charts, int documentRelationId)
        { 
            var documentRelationMapping =
                documentRelations.Descendants(P + "Relationship")
                                 .Where(x => x.Attribute("Type")?.Value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
                                          || x.Attribute("Type")?.Value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
                                 .Select(
                                     x => new
                                     {
                                         Id = x.Attribute("Id"),
                                         Type = x.Attribute("Type"),
                                         Target = x.Attribute("Target"),
                                         TargetMode = x.Attribute("TargetMode")
                                     })
                                 .OrderBy(x => x.Id.Value.ParseInt())
                                 .Select(
                                     (x, i) => new
                                     {
                                         oldId = x.Id,
                                         newId = new XAttribute("Id", $"rId{documentRelationId + i}"),
                                         x.Type,
                                         oldTarget = x.Target,
                                         newTarget = x.Target.Value.StartsWith("charts/") ? new XAttribute("Target", $"charts/chart{documentRelationId + i}.xml") : x.Target,
                                         x.TargetMode
                                     })
                                 .ToArray();

            ChartInformation[] chartMapping =
                documentRelationMapping
                    .Where(x => x.oldTarget.Value.StartsWith("charts/"))
                    .Select(
                        x => new
                        {
                            x.newId,
                            x.oldTarget,
                            x.newTarget
                        })
                    .OrderBy(x => x.newId.Value.ParseInt())
                    .Select(
                        x => new ChartInformation(x.newTarget.Value, charts.Single(y => y.Name == x.oldTarget.Value).Chart))
                    .Select(
                        x =>
                        {
                            x.Chart.Descendants(C + "externalData").Remove();
                            return x;
                        })
                    .ToArray();

            XElement modifiedDocument = 
                document.RemoveRsidAttributes();

            foreach (var map in documentRelationMapping)
            {
                modifiedDocument =
                    modifiedDocument.ChangeXAttributeValues(R + "id", map.oldId.Value, map.newId.Value);
            }

            XElement modifiedDocumentRelations =
                new XElement(
                    documentRelations.Name,
                    documentRelationMapping.Select(
                        x =>
                            new XElement(
                                P + "Relationship",
                                x.newId,
                                x.Type,
                                x.newTarget,
                                x.TargetMode)));

            XElement modifiedContentTypes =
                new XElement(
                    contentTypes.Name,
                    chartMapping.Select(
                        x =>
                            new XElement(T + "Override",
                                new XAttribute("PartName", $"/word/{x.Name}"),
                                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"))));

            return (modifiedDocument, modifiedDocumentRelations, modifiedContentTypes, chartMapping);
        }
    }
}