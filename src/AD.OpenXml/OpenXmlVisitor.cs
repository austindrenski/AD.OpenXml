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
        /// <summary>
        /// Represents the 'a:' prefix seen in the markup for chart[#].xml
        /// </summary>
        [NotNull] protected static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        /// <summary>
        /// Represents the 'c:' prefix seen in the markup for chart[#].xml
        /// </summary>
        [NotNull] protected static readonly XNamespace C = XNamespaces.OpenXmlDrawingmlChart;

        // TODO: move into AD.Xml
        /// <summary>
        /// Represents the 'dgm:' prefix seen in the markup for 'drawing' elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

        // TODO: move into AD.Xml
        /// <summary>
        /// Represents the 'm:' prefix seen in the markup for math elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

        // TODO: move into AD.Xml
        /// <summary>
        /// Represents the 'pic:' prefix seen in the markup for 'drawing' elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";

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

        // TODO: move into AD.Xml
        /// <summary>
        /// Represents the 'wps:' prefix seen in the markup for 'wsp' elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        /// <summary>
        /// Return an element when handling the default dispatch case if true; otherwise, false.
        /// </summary>
        private readonly bool _returnOnDefault;

        /// <summary>
        /// The mapping of <see cref="XName"/> to visit method used by <see cref="VisitElement"/>.
        /// </summary>
        protected readonly IDictionary<XName, Func<XElement, XObject>> VisitLookup;

        /// <summary>
        /// Initializes an <see cref="OpenXmlVisitor"/>.
        /// </summary>
        /// <param name="returnOnDefault">
        /// True if an element should be returned when handling the default dispatch case.
        /// </param>
        protected OpenXmlVisitor(bool returnOnDefault)
        {
            _returnOnDefault = returnOnDefault;

            VisitLookup =
                new Dictionary<XName, Func<XElement, XObject>>
                {
                    // @formatter:off
                    [A   + "graphic"]     = VisitGraphic,
                    [A   + "graphicData"] = VisitGraphicData,
                    [C   + "areaChart"]   = VisitAreaChart,
                    [C   + "barChart"]    = VisitBarChart,
                    [C   + "chart"]       = VisitChart,
                    [C   + "lineChart"]   = VisitLineChart,
                    [C   + "pieChart"]    = VisitPieChart,
                    [C   + "plotArea"]    = VisitPlotArea,
                    [DGM + "relIds"]      = VisitDiagram,
                    [M   + "oMath"]       = VisitMath,
                    [PIC + "pic"]         = VisitPicture,
                    [W   + "body"]        = VisitBody,
                    [W   + "document"]    = VisitDocument,
                    [W   + "drawing"]     = VisitDrawing,
                    [W   + "footnote"]    = VisitFootnote,
                    [W   + "footnotes"]   = VisitFootnotes,
                    [W   + "p"]           = VisitParagraph,
                    [W   + "r"]           = VisitRun,
                    [W   + "tbl"]         = VisitTable,
                    [W   + "tc"]          = VisitTableCell,
                    [W   + "tr"]          = VisitTableRow,
                    [WP  + "anchor"]      = VisitAnchor,
                    [WP  + "docPr"]       = VisitDocumentProperty,
                    [WP  + "inline"]      = VisitInline,
                    [WPS + "wsp"]         = VisitShape
                    // @formatter:on
                };
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="anchor">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAnchor([NotNull] XElement anchor)
        {
            if (anchor is null)
            {
                throw new ArgumentNullException(nameof(anchor));
            }

            return base.VisitElement(anchor);
        }

        /// <summary>
        /// Visits the area chart node.
        /// </summary>
        /// <param name="areaChart">
        /// The area chart to visit.
        /// </param>
        /// <returns>
        /// The reconstructed area chart.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAreaChart([NotNull] XElement areaChart)
        {
            if (areaChart is null)
            {
                throw new ArgumentNullException(nameof(areaChart));
            }

            return base.VisitElement(areaChart);
        }

        /// <summary>
        /// Visits the bar chart node.
        /// </summary>
        /// <param name="barChart">
        /// The bar chart to visit.
        /// </param>
        /// <returns>
        /// The reconstructed bar chart.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitBarChart([NotNull] XElement barChart)
        {
            if (barChart is null)
            {
                throw new ArgumentNullException(nameof(barChart));
            }

            return base.VisitElement(barChart);
        }

        /// <summary>
        /// Visits the body node.
        /// </summary>
        /// <param name="body">
        /// The body to visit.
        /// </param>
        /// <returns>
        /// The reconstructed attribute.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitBody([NotNull] XElement body)
        {
            if (body is null)
            {
                throw new ArgumentNullException(nameof(body));
            }

            return base.VisitElement(body);
        }

        /// <summary>
        /// Visits the chart node.
        /// </summary>
        /// <param name="chart">
        /// The chart to visit.
        /// </param>
        /// <returns>
        /// The reconstructed chart.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitChart([NotNull] XElement chart)
        {
            if (chart is null)
            {
                throw new ArgumentNullException(nameof(chart));
            }

            return base.VisitElement(chart);
        }

        /// <summary>
        /// Visits the document node.
        /// </summary>
        /// <param name="document">
        /// The document node to visit.
        /// </param>
        /// <returns>
        /// The reconstructed document.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDocument([NotNull] XElement document)
        {
            if (document is null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            return base.VisitElement(document);
        }

        /// <summary>
        /// Visits the document property node.
        /// </summary>
        /// <param name="docPr">
        /// The document property node.
        /// </param>
        /// <returns>
        /// The reconstructed document property.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDocumentProperty([NotNull] XElement docPr)
        {
            if (docPr is null)
            {
                throw new ArgumentNullException(nameof(docPr));
            }

            return base.VisitElement(docPr);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="diagram">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDiagram([NotNull] XElement diagram)
        {
            if (diagram is null)
            {
                throw new ArgumentNullException(nameof(diagram));
            }

            return base.VisitElement(diagram);
        }

        /// <summary>
        /// Visits the drawing node.
        /// </summary>
        /// <param name="drawing">
        /// The drawing node to visit.
        /// </param>
        /// <returns>
        /// The reconstructed drawing.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDrawing([NotNull] XElement drawing)
        {
            if (drawing is null)
            {
                throw new ArgumentNullException(nameof(drawing));
            }

            return base.VisitElement(drawing);
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitElement(XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            return
                VisitLookup.TryGetValue(element.Name, out Func<XElement, XObject> visit)
                    ? visit(element)
                    : _returnOnDefault
                        ? base.VisitElement(element)
                        : null;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="footnote">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFootnote([NotNull] XElement footnote)
        {
            if (footnote is null)
            {
                throw new ArgumentNullException(nameof(footnote));
            }

            return base.VisitElement(footnote);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="footnotes">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFootnotes([NotNull] XElement footnotes)
        {
            if (footnotes is null)
            {
                throw new ArgumentNullException(nameof(footnotes));
            }

            return base.VisitElement(footnotes);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="graphic">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitGraphic([NotNull] XElement graphic)
        {
            if (graphic is null)
            {
                throw new ArgumentNullException(nameof(graphic));
            }

            return base.VisitElement(graphic);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="graphicData">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitGraphicData([NotNull] XElement graphicData)
        {
            if (graphicData is null)
            {
                throw new ArgumentNullException(nameof(graphicData));
            }

            return base.VisitElement(graphicData);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="inline">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitInline([NotNull] XElement inline)
        {
            if (inline is null)
            {
                throw new ArgumentNullException(nameof(inline));
            }

            return base.VisitElement(inline);
        }

        /// <summary>
        /// Visits the line chart node.
        /// </summary>
        /// <param name="lineChart">
        /// The line chart to visit.
        /// </param>
        /// <returns>
        /// The visited line chart.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitLineChart([NotNull] XElement lineChart)
        {
            if (lineChart is null)
            {
                throw new ArgumentNullException(nameof(lineChart));
            }

            return base.VisitElement(lineChart);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="paragraph">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitParagraph([NotNull] XElement paragraph)
        {
            if (paragraph is null)
            {
                throw new ArgumentNullException(nameof(paragraph));
            }

            return base.VisitElement(paragraph);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="math">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitMath([NotNull] XElement math)
        {
            if (math is null)
            {
                throw new ArgumentNullException(nameof(math));
            }

            return base.VisitElement(math);
        }

        /// <summary>
        /// Visits the picture node.
        /// </summary>
        /// <param name="picture">
        /// The picture to visit.
        /// </param>
        /// <returns>
        /// The reconstructed picture.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitPicture([NotNull] XElement picture)
        {
            if (picture is null)
            {
                throw new ArgumentNullException(nameof(picture));
            }

            return base.VisitElement(picture);
        }

        /// <summary>
        /// Visits the pie chart node.
        /// </summary>
        /// <param name="pieChart">
        /// The pie chart to visit.
        /// </param>
        /// <returns>
        /// The visited pie chart.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitPieChart([NotNull] XElement pieChart)
        {
            if (pieChart is null)
            {
                throw new ArgumentNullException(nameof(pieChart));
            }

            return base.VisitElement(pieChart);
        }

        /// <summary>
        /// Visits the plot area node.
        /// </summary>
        /// <param name="plotArea">
        /// The plot area to visit.
        /// </param>
        /// <returns>
        /// The visited plot area.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitPlotArea([NotNull] XElement plotArea)
        {
            if (plotArea is null)
            {
                throw new ArgumentNullException(nameof(plotArea));
            }

            return base.VisitElement(plotArea);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="run">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitRun([NotNull] XElement run)
        {
            if (run is null)
            {
                throw new ArgumentNullException(nameof(run));
            }

            return base.VisitElement(run);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="shape">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        private XObject VisitShape([NotNull] XElement shape)
        {
            if (shape is null)
            {
                throw new ArgumentNullException(nameof(shape));
            }

            return base.VisitElement(shape);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="table">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitTable([NotNull] XElement table)
        {
            if (table is null)
            {
                throw new ArgumentNullException(nameof(table));
            }

            return base.VisitElement(table);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="cell">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitTableCell([NotNull] XElement cell)
        {
            if (cell is null)
            {
                throw new ArgumentNullException(nameof(cell));
            }

            return base.VisitElement(cell);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="row">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitTableRow([NotNull] XElement row)
        {
            if (row is null)
            {
                throw new ArgumentNullException(nameof(row));
            }

            return base.VisitElement(row);
        }
    }
}