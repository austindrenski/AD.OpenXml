using System;
using System.Collections.Generic;
using System.Linq;
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
        /// xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main".
        /// </summary>
        [NotNull] private static readonly XNamespace A = XNamespaces.OpenXmlDrawingmlMain;

        /// <summary>
        /// Return an element when handling the default dispatch case if true; otherwise, false.
        /// </summary>
        private readonly bool _returnOnDefault;

        /// <summary>
        /// The mapping of chart id to node.
        /// </summary>
        [NotNull]
        protected abstract IDictionary<string, XElement> Charts { get; set; }

        /// <summary>
        /// The mapping of image id to data.
        /// </summary>
        protected abstract IDictionary<string, (string mime, string description, string base64)> Images { get; set; }

        /// <summary>
        ///
        /// </summary>
        /// <param name="returnOnDefault">
        ///
        /// </param>
        protected OpenXmlVisitor(bool returnOnDefault)
        {
            _returnOnDefault = returnOnDefault;
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

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitElement(XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            switch (element)
            {
                case XElement e when e.Name.LocalName == "anchor":
                {
                    return VisitAnchor(e);
                }
                case XElement e when e.Name.LocalName == "areaChart":
                {
                    return VisitAreaChart(e);
                }
                case XElement e when e.Name.LocalName == "barChart":
                {
                    return VisitBarChart(e);
                }
                case XElement e when e.Name.LocalName == "body":
                {
                    return VisitBody(e);
                }
                case XElement e when e.Name.LocalName == "chart":
                {
                    return VisitChart(e);
                }
                case XElement e when e.Name.LocalName == "drawing":
                {
                    return VisitDrawing(e);
                }
                case XElement e when e.Name.LocalName == "footnote":
                {
                    return VisitFootnote(e);
                }
                case XElement e when e.Name.LocalName == "graphic":
                {
                    return VisitGraphic(e);
                }
                case XElement e when e.Name.LocalName == "graphicData":
                {
                    return VisitGraphicData(e);
                }
                case XElement e when e.Name.LocalName == "inline":
                {
                    return VisitInline(e);
                }
                case XElement e when e.Name.LocalName == "lineChart":
                {
                    return VisitLineChart(e);
                }
                case XElement e when e.Name.LocalName == "p":
                {
                    return VisitParagraph(e);
                }
                case XElement e when e.Name.LocalName == "pic":
                {
                    return VisitPicture(e);
                }
                case XElement e when e.Name.LocalName == "pieChart":
                {
                    return VisitPieChart(e);
                }
                case XElement e when e.Name.LocalName == "r":
                {
                    return VisitRun(e);
                }
                case XElement e when e.Name.LocalName == "relIds":
                {
                    return VisitDiagram(e);
                }
                case XElement e when e.Name.LocalName == "tbl":
                {
                    return VisitTable(e);
                }
                case XElement e when e.Name.LocalName == "tr":
                {
                    return VisitTableRow(e);
                }
                case XElement e when e.Name.LocalName == "tc":
                {
                    return VisitTableCell(e);
                }
                case XElement e when e.Name.LocalName == "wsp":
                {
                    return VisitShape(e);
                }
                default:
                {
                    return _returnOnDefault ? base.VisitElement(element) : null;
                }
            }
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