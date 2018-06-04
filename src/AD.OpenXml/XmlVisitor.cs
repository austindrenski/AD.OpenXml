using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

// ReSharper disable VirtualMemberNeverOverridden.Global

namespace AD.OpenXml
{
    /// <summary>
    /// Defines an XML visitor that clones an existing <see cref="XObject"/>.
    /// </summary>
    [PublicAPI]
    public abstract class XmlVisitor
    {
        /// <summary>
        /// The "data-liftable" attribute.
        /// </summary>
        [NotNull] protected const string Liftable = "data-liftable";

        /// <summary>
        /// Visits the <see cref="XObject"/>.
        /// </summary>
        /// <param name="xObject">
        /// The <see cref="XObject"/> to visit.
        /// </param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        public XObject Visit([CanBeNull] XObject xObject)
        {
            switch (xObject)
            {
                case null:
                {
                    return null;
                }
                case XAttribute a:
                {
                    return VisitAttribute(a);
                }
                case XElement e:
                {
                    return VisitElement(e);
                }
                case XText t:
                {
                    return VisitText(t);
                }
                case XDocument _:
                {
                    throw new NotImplementedException();
                }
                case XContainer _:
                {
                    throw new NotImplementedException();
                }
                case XComment _:
                {
                    throw new NotImplementedException();
                }
                case XDocumentType _:
                {
                    throw new NotImplementedException();
                }
                case XProcessingInstruction _:
                {
                    throw new NotImplementedException();
                }
                case XNode _:
                {
                    throw new NotImplementedException();
                }
                default:
                {
                    return VisitObject(xObject);
                }
            }
        }

        /// <summary>
        /// Visits an <see cref="XObject"/> collection.
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
        protected virtual IEnumerable<XObject> Visit([NotNull] [ItemCanBeNull] IEnumerable<XObject> source)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            foreach (XObject item in source)
            {
                yield return Visit(item);
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
            => text is null
                   ? null
                   : Visit(new XText(text));

        /// <summary>
        /// Visits an <see cref="XAttribute"/>.
        /// </summary>
        /// <param name="attribute">
        /// The <see cref="XAttribute"/> to visit.
        /// </param>
        /// <returns>
        /// The visited <see cref="XAttribute"/>.
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

            return
                new XAttribute(
                    VisitName(attribute.Name),
                    attribute.Value);
        }

        /// <summary>
        /// Visits an <see cref="XElement"/>.
        /// </summary>
        /// <param name="element">
        /// The <see cref="XElement"/> to visit.
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
                    VisitName(element.Name),
                    Visit(element.Attributes()),
                    Visit(element.Nodes()));
        }

        /// <summary>
        /// Visits an <see cref="XName"/>.
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
        protected virtual XName VisitName([NotNull] XName name)
        {
            if (name is null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            return name;
        }

        /// <summary>
        /// Visits an <see cref="XObject"/> that was unhandled by known derived type methods.
        /// </summary>
        /// <param name="xObject">
        /// The <see cref="XObject"/> to visit.
        /// </param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected XObject VisitObject([CanBeNull] XObject xObject) => xObject;

        /// <summary>
        /// Visits an <see cref="XText"/>.
        /// </summary>
        /// <param name="text">
        /// The <see cref="XText"/> to visit.
        /// </param>
        /// <returns>
        /// The visited <see cref="XText"/>.
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

            return new XText(text.Value);
        }

        /// <summary>
        /// Lifts a singleton <see cref="XObject"/>.
        /// </summary>
        /// <param name="xObject">
        /// The <see cref="XObject"/> to visit and lift.
        /// </param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected static XObject LiftSingleton([CanBeNull] XObject xObject)
            => xObject is XContainer container && container.Nodes().Count() <= 1
                   ? container.FirstNode
                   : xObject;

        /// <summary>
        /// Yields the <see cref="XObject"/> or the children of the <see cref="XObject"/> if the "data-liftable" attribute is present.
        /// </summary>
        /// <param name="xObject">
        /// The <see cref="XObject"/> to visit.
        /// </param>
        /// <returns>
        /// The <see cref="XObject"/> or the children of the <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [NotNull]
        [ItemCanBeNull]
        protected static IEnumerable<XObject> Lift([CanBeNull] XObject xObject)
        {
            if (!(xObject is XElement e))
            {
                yield return xObject;

                yield break;
            }

            if (e.Attribute(Liftable) is null)
            {
                yield return new XElement(e.Name, e.Attributes(), e.Nodes().Select(Lift));

                yield break;
            }

            foreach (XObject item in e.Elements().SelectMany(Lift))
            {
                yield return item;
            }
        }

        /// <summary>
        /// Yields the <see cref="XObject"/> or the children of the <see cref="XObject"/> if the "data-liftable" attribute is present.
        /// </summary>
        /// <param name="source">
        /// The collection of <see cref="XObject"/> to visit.
        /// </param>
        /// <returns>
        /// The <see cref="XObject"/> or the children of the <see cref="XObject"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        [ItemCanBeNull]
        protected static IEnumerable<XObject> Lift([NotNull] IEnumerable<XObject> source)
        {
            if (source is null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            foreach (XObject item in source.SelectMany(Lift))
            {
                yield return item;
            }
        }

        /// <summary>
        /// Constructs a liftable div element.
        /// </summary>
        /// <param name="element">
        /// The <see cref="XElement"/> from which nodes can be lifted.
        /// </param>
        /// <returns>
        /// An <see cref="XObject"/> representing the lift operation.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected XObject LiftableHelper([CanBeNull] XElement element)
            => element is null
                   ? null
                   : new XElement("div",
                       new XAttribute(Liftable, $"from-{element.Name.LocalName}"),
                       Visit(element.Elements()));

        ///  <summary>
        ///
        ///  </summary>
        ///  <param name="current">
        ///
        ///  </param>
        /// <param name="cast">
        ///
        /// </param>
        /// <param name="predicates">
        ///
        ///  </param>
        ///  <returns>
        ///
        ///  </returns>
        [Pure]
        [NotNull]
        [ItemNotNull]
        protected IEnumerable<T> NextWhile<T>([NotNull] XNode current, Func<XNode, T> cast, [NotNull] [ItemNotNull] params Func<T, bool>[] predicates) where T : class
        {
            if (current is null)
            {
                throw new ArgumentNullException(nameof(current));
            }

            if (predicates.Any(x => x is null))
            {
                throw new ArgumentNullException(nameof(predicates));
            }

            for (XNode next = current.NextNode; next != null; next = next.NextNode)
            {
                T item = cast(next);

                if (item is null)
                {
                    yield break;
                }

                for (int i = 0; i < predicates.Length; i++)
                {
                    if (!predicates[i](item))
                    {
                        yield break;
                    }
                }

                yield return item;
            }
        }
    }
}