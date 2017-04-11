using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.IO;
using AD.OpenXml.Elements;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    /// <summary>
    /// This class serves as a container to encapsulate XML components of a Word document.
    /// </summary>
    [PublicAPI]
    public class OpenXmlContainer
    {
        private static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// word/document.xml
        /// </summary>
        private readonly XElement _sourceDocument;

        private XElement _document;

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        private readonly XElement _sourceContentTypes;

        private XElement _contentTypes;
        
        /// <summary>
        /// word/footnotes.xml
        /// </summary>
        private readonly XElement _sourceFootnotes;

        private XElement _footnotes;
        
        /// <summary>
        /// word/_rels/document.xml.rels
        /// </summary>
        private readonly XElement _sourcedocumentRelations;

        private XElement _documentRelations;
        
        private int CurrentFootnoteId
        {
            get
            {
                return _footnotes.Elements(W + "footnote")
                                 .Attributes(W + "id")
                                 .Select(x => int.Parse(x.Value))
                                 .Max();
            }
        }

        /// <summary>
        /// Initializes an <see cref="OpenXmlContainer"/> by reading document parts into memory.
        /// </summary>
        /// <param name="contentFilePath"></param>
        public OpenXmlContainer(DocxFilePath contentFilePath)
        {
            _sourceDocument = contentFilePath.ReadAsXml("word/document.xml");
            _document = _sourceDocument.Clone();

            _sourceContentTypes = contentFilePath.ReadAsXml("[Content_Types].xml");
            _contentTypes = _sourceContentTypes.Clone();

            _sourceFootnotes = contentFilePath.ReadAsXml("word/footnotes.xml");
            _footnotes = _sourceFootnotes.Clone();

            _sourceDocumentRelations = contentFilePath.ReadAsXml("word/document.xml.rels");
            _documentRelations = _sourcedocumentRelations.Clone()
        }

        /// <summary>
        /// Saves any modifications to <paramref name="resultPath"/>. This operation will overwrite any existing content for the modified parts.
        /// </summary>
        /// <param name="resultPath">The path to which modified parts are written.</param>
        public void Save([NotNull] DocxFilePath resultPath)
        {
            if (!XNode.DeepEquals(_sourceDocument, _document))
            {
                _document.WriteInto(resultPath, "word/document.xml");
            }
            if (!XNode.DeepEquals(_sourceContentTypes, _footnotes))
            {
                _footnotes.WriteInto(resultPath, "word/footnotes.xml");
            }
            if (!XNode.DeepEquals(_sourceContentTypes, _contentTypes))
            {
                _contentTypes.WriteInto(resultPath, "[Content_Types].xml");
            }
            if (!XNode.DeepEquals(_documentRelations, _documentRelations))
            {
                _documentRelations.WriteInto(resultPath, "word/_rels/document.xml.rels");
            }
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="source">The file from which content is copied.</param>
        public void MergeDocuments([NotNull] DocxFilePath source)
        {
            MergeContentFrom(source);
            MergeFootnotesFrom(source);
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="source">The file from which content is copied.</param>
        private void MergeContentFrom([NotNull] DocxFilePath source)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            IEnumerable<XElement> sourceContent =
                source.ReadAsXml("word/document.xml")
                      .Process508From()
                      .Elements()
                      .Single()
                      .Elements();

            _document.Elements()
                     .First()
                     .Add(sourceContent);
        }

        /// <summary>
        /// Merges the source document into the result document.
        /// </summary>
        /// <param name="source">The file from which content is copied.</param>
        private void MergeFootnotesFrom([NotNull] DocxFilePath source)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            IEnumerable<XElement> sourceFootnotes =
                source.ReadAsXml("word/footnotes.xml")
                      .Elements(W + "footnote")
                      .RemoveRsidAttributes()
                      .ToArray();

            sourceFootnotes.Descendants(W + "p")
                           .Attributes()
                           .Remove();

            sourceFootnotes.Descendants(W + "hyperlink")
                           .Remove();

            var footnoteMapping =
                sourceFootnotes.Attributes(W + "id")
                               .Select(x => x.Value)
                               .Select(int.Parse)
                               .Where(x => x > 0)
                               .OrderByDescending(x => x)
                               .Select(
                                   x => new
                                   {
                                       oldId = $"{x}",
                                       newId = $"{CurrentFootnoteId + x}"
                                   });

            foreach (var map in footnoteMapping)
            {
                _document = 
                    _document.ChangeXAttributeValues(W + "footnoteReference", W + "Id", map.oldId, map.newId);

                sourceFootnotes =
                    sourceFootnotes.ChangeXAttributeValues(W + "footnote", W + "id", map.oldId, map.newId)
                                   .ToArray();
            }

            _footnotes.Add(sourceFootnotes);
        }
    }
}