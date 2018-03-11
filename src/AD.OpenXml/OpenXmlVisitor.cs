using System;
using System.Collections.Generic;
using System.Xml.Linq;
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
        /// The attributes that may be returned.
        /// </summary>
        [NotNull]
        [ItemNotNull]
        protected abstract ISet<XName> SupportedAttributes { get; }

        /// <summary>
        /// The elements that may be returned.
        /// </summary>
        [NotNull]
        [ItemNotNull]
        protected abstract ISet<XName> SupportedElements { get; }

        /// <summary>
        /// The mapping between OpenXML names and supported names.
        /// </summary>
        [NotNull]
        protected abstract IDictionary<XName, XName> Renames { get; }

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

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitAttribute(XAttribute attribute)
        {
            if (attribute is null)
            {
                throw new ArgumentNullException(nameof(attribute));
            }

            XName name = VisitName(attribute.Name);

            return SupportedAttributes.Contains(name) ? new XAttribute(name, attribute.Value) : null;
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
        ///
        /// </summary>
        /// <param name="drawing">
        ///
        /// </param>
        /// <returns>
        ///
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

            switch (element)
            {
                case XElement e when e.Name.LocalName == "body":
                {
                    return VisitBody(e);
                }
                case XElement e when e.Name.LocalName == "areaChart":
                {
                    return VisitAreaChart(e);
                }
                case XElement e when e.Name.LocalName == "barChart":
                {
                    return VisitBarChart(e);
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
                case XElement e when e.Name.LocalName == "p":
                {
                    return VisitParagraph(e);
                }
                case XElement e when e.Name.LocalName == "pic":
                {
                    return VisitPicture(e);
                }
                case XElement e when e.Name.LocalName == "r":
                {
                    return VisitRun(e);
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
                default:
                {
                    return base.VisitElement(element);
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

        /// <inheritdoc />
        [Pure]
        protected override XName VisitName(XName name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return Renames.TryGetValue(name.LocalName, out XName result) ? result : name.LocalName;
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