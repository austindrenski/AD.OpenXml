using System;
using System.Collections.Generic;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

// ReSharper disable VirtualMemberNeverOverridden.Global
namespace AD.OpenXml
{
    /// <inheritdoc />
    /// <summary>
    /// Defines an abstract base class to visit and rewrite a minimalist OpenXML node.
    /// </summary>
    [PublicAPI]
    public abstract class OpenXmlVisitor : XmlVisitor
    {
        #region Namespaces

        /// <summary>
        /// Represents the 'a:' prefix seen in the markup for chart[#].xml
        /// </summary>
        [NotNull] protected static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        /// <summary>
        /// Represents the 'c:' prefix seen in the markup for chart[#].xml
        /// </summary>
        [NotNull] protected static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        /// <summary>
        /// Represents the 'dgm:' prefix seen in the markup for 'drawing' elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace DGM = XNamespaces.OpenXmlDrawingmlDiagram;

        /// <summary>
        /// Represents the 'm:' prefix seen in the markup for math elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace M = XNamespaces.OpenXmlMath;

        // TODO: move to AD.Xml
        /// <summary>
        /// Represents the 'mc:' prefix seen in the markup for compatibility blocks.
        /// </summary>
        [NotNull] protected static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        // TODO: move to AD.Xml
        /// <summary>
        /// Represents the 'o:' prefix seen in the markup for OLE elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";

        /// <summary>
        /// Represents the 'pic:' prefix seen in the markup for 'drawing' elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace PIC = XNamespaces.OpenXmlDrawingmlPicture;

        /// <summary>
        /// Represents the 'r:' prefix seen in the markup of document.xml.
        /// </summary>
        [NotNull] protected static readonly XNamespace R = XNamespaces.OpenXmlOfficeDocumentRelationships;

        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull] protected static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// Represents the 'wp:' prefix seen in the markup for 'drawing' elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace WP = XNamespaces.OpenXmlDrawingmlWordprocessingDrawing;

        /// <summary>
        /// Represents the 'wps:' prefix seen in the markup for 'wsp' elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace WPS = XNamespaces.OpenXmlWordprocessingShape;

        #endregion

        #region Fields

        /// <summary>
        /// True if the base method should be called when handling the default dispatch case.
        /// </summary>
        private readonly bool _allowBaseMethod;

        /// <summary>
        /// The mapping of <see cref="XName"/> to visit method used by <see cref="VisitElement"/>.
        /// </summary>
        protected readonly IDictionary<XName, Func<XElement, XObject>> VisitLookup;

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes an <see cref="OpenXmlVisitor"/>.
        /// </summary>
        /// <param name="allowBaseMethod">True if the base method should be called when handling the default dispatch case.</param>
        protected OpenXmlVisitor(bool allowBaseMethod)
        {
            _allowBaseMethod = allowBaseMethod;

            VisitLookup =
                new Dictionary<XName, Func<XElement, XObject>>
                {
                    // @formatter:off
                    [A   + "graphic"]          = VisitGraphic,
                    [A   + "graphicData"]      = VisitGraphicData,
                    [C   + "areaChart"]        = VisitAreaChart,
                    [C   + "barChart"]         = VisitBarChart,
                    [C   + "chart"]            = VisitChart,
                    [C   + "lineChart"]        = VisitLineChart,
                    [C   + "pieChart"]         = VisitPieChart,
                    [C   + "plotArea"]         = VisitPlotArea,
                    [DGM + "relIds"]           = VisitDiagram,
                    [M   + "oMath"]            = VisitMath,
                    [M   + "oMathPara"]        = VisitMathParagraph,
                    [O   + "OLEObject"]        = VisitOLEObject,
                    [PIC + "pic"]              = VisitPicture,
                    [MC  + "AlternateContent"] = VisitAlternateContent,
                    [W   + "body"]             = VisitBody,
                    [W   + "br"]               = VisitBreak,
                    [W   + "document"]         = VisitDocument,
                    [W   + "drawing"]          = VisitDrawing,
                    [W   + "footnote"]         = VisitFootnote,
                    [W   + "footnotes"]        = VisitFootnotes,
                    [W   + "hyperlink"]        = VisitHyperlink,
                    [W   + "object"]           = VisitEmbedded,
                    [W   + "p"]                = VisitParagraph,
                    [W   + "pPr"]              = VisitParagraphProperties,
                    [W   + "pStyle"]           = VisitParagraphStyle,
                    [W   + "r"]                = VisitRun,
                    [W   + "t"]                = VisitText,
                    [W   + "tbl"]              = VisitTable,
                    [W   + "tc"]               = VisitTableCell,
                    [W   + "tr"]               = VisitTableRow,
                    [WP  + "anchor"]           = VisitAnchor,
                    [WP  + "docPr"]            = VisitDocumentProperty,
                    [WP  + "inline"]           = VisitInline,
                    [WPS + "wsp"]              = VisitShape
                    // @formatter:on
                };
        }

        #endregion

        #region Visits

        /// <summary>
        ///
        /// </summary>
        /// <param name="alternate"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAlternateContent([NotNull] XElement alternate) => base.VisitElement(alternate);

        /// <summary>
        ///
        /// </summary>
        /// <param name="anchor"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAnchor([NotNull] XElement anchor) => base.VisitElement(anchor);

        /// <summary>
        /// Visits the area chart node.
        /// </summary>
        /// <param name="areaChart">The area chart to visit.</param>
        /// <returns>
        /// The reconstructed area chart.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAreaChart([NotNull] XElement areaChart) => base.VisitElement(areaChart);

        /// <summary>
        /// Visits the bar chart node.
        /// </summary>
        /// <param name="barChart">The bar chart to visit.</param>
        /// <returns>
        /// The reconstructed bar chart.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitBarChart([NotNull] XElement barChart) => base.VisitElement(barChart);

        /// <summary>
        /// Visits the break node.
        /// </summary>
        /// <param name="br">The break to visit.</param>
        /// <returns>
        /// The reconstructed break.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitBreak([NotNull] XElement br) => base.VisitElement(br);

        /// <summary>
        /// Visits the body node.
        /// </summary>
        /// <param name="body">The body to visit.</param>
        /// <returns>
        /// The reconstructed attribute.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitBody([NotNull] XElement body) => base.VisitElement(body);

        /// <summary>
        /// Visits the chart node.
        /// </summary>
        /// <param name="chart">The chart to visit.</param>
        /// <returns>
        /// The reconstructed chart.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitChart([NotNull] XElement chart) => base.VisitElement(chart);

        /// <summary>
        /// Visits the document node.
        /// </summary>
        /// <param name="document">The document node to visit.</param>
        /// <returns>
        /// The reconstructed document.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDocument([NotNull] XElement document) => base.VisitElement(document);

        /// <summary>
        /// Visits the document property node.
        /// </summary>
        /// <param name="docPr">The document property node.</param>
        /// <returns>
        /// The reconstructed document property.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDocumentProperty([NotNull] XElement docPr) => base.VisitElement(docPr);

        /// <summary>
        ///
        /// </summary>
        /// <param name="diagram"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDiagram([NotNull] XElement diagram) => base.VisitElement(diagram);

        /// <summary>
        /// Visits the drawing node.
        /// </summary>
        /// <param name="drawing">The drawing node to visit.</param>
        /// <returns>
        /// The reconstructed drawing.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDrawing([NotNull] XElement drawing) => base.VisitElement(drawing);

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitElement(XElement element)
            => VisitLookup.TryGetValue(element.Name, out Func<XElement, XObject> visit)
                   ? visit(element)
                   : _allowBaseMethod
                       ? base.VisitElement(element)
                       : null;

        /// <summary>
        ///
        /// </summary>
        /// <param name="embedded"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitEmbedded([NotNull] XElement embedded) => base.VisitElement(embedded);

        /// <summary>
        ///
        /// </summary>
        /// <param name="footnote"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFootnote([NotNull] XElement footnote) => base.VisitElement(footnote);

        /// <summary>
        ///
        /// </summary>
        /// <param name="footnotes"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFootnotes([NotNull] XElement footnotes) => base.VisitElement(footnotes);

        /// <summary>
        ///
        /// </summary>
        /// <param name="graphic"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitGraphic([NotNull] XElement graphic) => base.VisitElement(graphic);

        /// <summary>
        ///
        /// </summary>
        /// <param name="graphicData"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitGraphicData([NotNull] XElement graphicData) => base.VisitElement(graphicData);

        /// <summary>
        ///
        /// </summary>
        /// <param name="hyperlink"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitHyperlink([NotNull] XElement hyperlink) => base.VisitElement(hyperlink);

        /// <summary>
        ///
        /// </summary>
        /// <param name="inline"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitInline([NotNull] XElement inline) => base.VisitElement(inline);

        /// <summary>
        /// Visits the line chart node.
        /// </summary>
        /// <param name="lineChart">The line chart to visit.</param>
        /// <returns>
        /// The visited line chart.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitLineChart([NotNull] XElement lineChart) => base.VisitElement(lineChart);

        /// <summary>
        ///
        /// </summary>
        /// <param name="paragraph"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitParagraph([NotNull] XElement paragraph) => base.VisitElement(paragraph);

        /// <summary>
        ///
        /// </summary>
        /// <param name="properties"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitParagraphProperties([NotNull] XElement properties) => base.VisitElement(properties);

        /// <summary>
        ///
        /// </summary>
        /// <param name="style"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitParagraphStyle([NotNull] XElement style) => base.VisitElement(style);

        /// <summary>
        ///
        /// </summary>
        /// <param name="math"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitMath([NotNull] XElement math) => base.VisitElement(math);

        /// <summary>
        ///
        /// </summary>
        /// <param name="mathParagraph"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitMathParagraph([NotNull] XElement mathParagraph) => base.VisitElement(mathParagraph);

        /// <summary>
        ///
        /// </summary>
        /// <param name="oleObject"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitOLEObject([NotNull] XElement oleObject) => base.VisitElement(oleObject);

        /// <summary>
        /// Visits the picture node.
        /// </summary>
        /// <param name="picture">The picture to visit.</param>
        /// <returns>
        /// The reconstructed picture.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitPicture([NotNull] XElement picture) => base.VisitElement(picture);

        /// <summary>
        /// Visits the pie chart node.
        /// </summary>
        /// <param name="pieChart">The pie chart to visit.</param>
        /// <returns>
        /// The visited pie chart.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitPieChart([NotNull] XElement pieChart) => base.VisitElement(pieChart);

        /// <summary>
        /// Visits the plot area node.
        /// </summary>
        /// <param name="plotArea">The plot area to visit.</param>
        /// <returns>
        /// The visited plot area.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitPlotArea([NotNull] XElement plotArea) => base.VisitElement(plotArea);

        /// <summary>
        ///
        /// </summary>
        /// <param name="run"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitRun([NotNull] XElement run) => base.VisitElement(run);

        /// <summary>
        ///
        /// </summary>
        /// <param name="section"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSectionProperties([NotNull] XElement section) => base.VisitElement(section);

        /// <summary>
        ///
        /// </summary>
        /// <param name="shape"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        private XObject VisitShape([NotNull] XElement shape) => base.VisitElement(shape);

        /// <summary>
        ///
        /// </summary>
        /// <param name="table"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitTable([NotNull] XElement table) => base.VisitElement(table);

        /// <summary>
        ///
        /// </summary>
        /// <param name="cell"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitTableCell([NotNull] XElement cell) => base.VisitElement(cell);

        /// <summary>
        ///
        /// </summary>
        /// <param name="row"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitTableRow([NotNull] XElement row) => base.VisitElement(row);

        /// <summary>
        ///
        /// </summary>
        /// <param name="text"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitText([NotNull] XElement text) => base.VisitElement(text);

        #endregion
    }
}