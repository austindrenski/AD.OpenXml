using System;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
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
        [NotNull] private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

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
            XElement numbering = Execute(subject.Numbering.Clone());

            Result =
                new OpenXmlVisitor(
                    subject.ContentTypes,
                    subject.Document,
                    subject.DocumentRelations,
                    subject.Footnotes,
                    subject.FootnoteRelations,
                    subject.Styles,
                    numbering,
                    subject.Charts);
        }

        [Pure]
        [NotNull]
        private static XElement Execute([NotNull] XElement numbering)
        {
            if (numbering is null)
            {
                throw new ArgumentNullException(nameof(numbering));
            }




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

            return XElement.Parse(Resources.Numbering);
        }
    }
}
