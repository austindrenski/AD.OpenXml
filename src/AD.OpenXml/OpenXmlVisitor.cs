using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

// ReSharper disable VirtualMemberNeverOverridden.Global

namespace AD.OpenXml
{
    /// <summary>
    /// Defines an abstract base class to visit and rewrite an OpenXML node.
    /// </summary>
    [PublicAPI]
    public abstract class OpenXmlVisitor
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
        /// Visits the node.
        /// </summary>
        /// <param name="xObject">
        /// The XML object to visit.
        /// </param>
        /// <returns>
        /// The visited node.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject Visit([CanBeNull] XObject xObject)
        {
            switch (xObject)
            {
                case XAttribute a:
                {
                    return VisitAttribute(a);
                }
                case XElement e when e.Name.LocalName == "body":
                {
                    return VisitBody(e);
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
                case XElement e when e.Name.LocalName == "r" && (e.Next()?.Elements()?.Any(x => x.Name.LocalName == "footnoteReference") ?? false):
                {
                    // OpenXML places footnote references in a trailing run.
                    // This provides an opportunity to catch the reference before the content.
                    // By default, this returns null.
                    // If you override this, then you must override VisitFootnoteReference.
                    return VisitFootnoteReferenceEarly(e);
                }
                case XElement e when e.Name.LocalName == "r" && e.Elements().Any(x => x.Name.LocalName == "footnoteReference"):
                {
                    // If VisitFootnoteReferenceEarly was overridden, you must override this too..
                    return VisitFootnoteReference(e);
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
                case XElement e:
                {
                    return VisitElement(e);
                }
                case XText t:
                {
                    return VisitText(t);
                }
                default:
                {
                    return xObject;
                }
            }
        }

        /// <summary>
        /// Visits the text as an <see cref="XText"/> node.
        /// </summary>
        /// <param name="text">
        /// The text to visit.
        /// </param>
        /// <returns>
        /// A visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject Visit([CanBeNull] string text)
        {
            return Visit(new XText(text));
        }

        /// <summary>
        /// Visits the <see cref="XObject"/> collection.
        /// </summary>
        /// <param name="source">
        /// The <see cref="XObject"/> collection to visit.
        /// </param>
        /// <returns>
        /// A visited <see cref="XObject"/> collection.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        [ItemCanBeNull]
        protected virtual IEnumerable<XObject> Visit([CanBeNull] [ItemCanBeNull] IEnumerable<XObject> source)
        {
            if (source is null)
            {
                yield break;
            }

            foreach (XObject item in source)
            {
                yield return Visit(item);
            }
        }

        /// <summary>
        /// Visits the <see cref="XName"/> for renaming and localization.
        /// </summary>
        /// <param name="name">
        /// The name to visit.
        /// </param>
        /// <returns>
        /// A visited <see cref="XName"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [NotNull]
        protected virtual XName Visit([NotNull] XName name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return Renames.TryGetValue(name.LocalName, out XName result) ? result : name.LocalName;
        }

        /// <summary>
        /// Reconstructs the attribute with only the local name.
        /// </summary>
        /// <param name="attribute">
        /// The attribute to rename.
        /// </param>
        /// <returns>
        /// The reconstructed attribute.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAttribute([NotNull] XAttribute attribute)
        {
            if (attribute is null)
            {
                throw new ArgumentNullException(nameof(attribute));
            }

            XName name = Visit(attribute.Name);

            return SupportedAttributes.Contains(name) ? new XAttribute(name, attribute.Value) : null;
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

            return VisitElement(body);
        }

        /// <summary>
        /// Visits an <see cref="XElement"/>.
        /// </summary>
        /// <param name="element">
        /// A generic <see cref="XElement"/>.
        /// </param>
        /// <returns>
        /// The visited <see cref="XElement"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitElement([NotNull] XElement element)
        {
            if (element is null)
            {
                throw new ArgumentNullException(nameof(element));
            }

            return
                new XElement(
                    Visit(element.Name),
                    Visit(element.Attributes()),
                    Visit(element.Nodes()));
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

            return VisitElement(footnote);
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
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFootnoteReference([NotNull] XElement run)
        {
            if (run is null)
            {
                throw new ArgumentNullException(nameof(run));
            }

            return VisitElement(run);
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
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFootnoteReferenceEarly([NotNull] XElement run)
        {
            if (run is null)
            {
                throw new ArgumentNullException(nameof(run));
            }

            return null;
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

            return VisitElement(drawing);
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

            return VisitElement(paragraph);
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

            return VisitElement(run);
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

            return VisitElement(table);
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

            return VisitElement(cell);
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

            return VisitElement(row);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="text">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitText([NotNull] XText text)
        {
            if (text is null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            return new XText(text);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="xObject">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject LiftSingleton([CanBeNull] XObject xObject)
        {
            if (xObject is XContainer container && container.Nodes().Count() <= 1)
            {
                return container.FirstNode;
            }

            return xObject;
        }
    }
}