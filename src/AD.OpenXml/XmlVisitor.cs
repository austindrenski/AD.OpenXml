using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
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
        [NotNull] private const string Liftable = "liftable";

        #region Main

        /// <summary>
        /// Visits the <see cref="XObject"/>.
        /// This method can dispatch to <see cref="VisitObject"/>.
        /// </summary>
        /// <param name="xObject">The <see cref="XObject"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        public XObject Visit([CanBeNull] XObject xObject) => xObject is null ? null : VisitObject(xObject);

        /// <summary>
        /// Visits an <see cref="XObject"/> collection.
        /// This method can dispatch to <see cref="Visit(XObject)"/>.
        /// </summary>
        /// <param name="source">The <see cref="XObject"/> collection to visit.</param>
        /// <param name="other">Additional <see cref="XObject"/> content to visit.</param>
        /// <returns>
        /// A visited <see cref="XObject"/> collection.
        /// </returns>
        [Pure]
        [NotNull]
        [LinqTunnel]
        [ItemNotNull]
        [CollectionAccess(CollectionAccessType.Read)]
        protected IEnumerable<XObject> Visit([CanBeNull] [ItemCanBeNull] IEnumerable<XObject> source, [CanBeNull] [ItemCanBeNull] params XObject[] other)
        {
            if (source is null)
                yield break;

            foreach (XObject item in source)
            {
                if (Visit(item) is XObject result)
                    yield return result;
            }

            if (other is null)
                yield break;

            foreach (XObject item in other)
            {
                if (Visit(item) is XObject result)
                    yield return result;
            }
        }

        #endregion

        #region Visits

        /// <summary>
        /// Visits an <see cref="XAttribute"/>.
        /// </summary>
        /// <param name="attribute">The <see cref="XAttribute"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAttribute([NotNull] XAttribute attribute)
            => new XAttribute(VisitName(attribute.Name), attribute.Value);

        /// <summary>
        /// Visits an <see cref="XCData"/>.
        /// </summary>
        /// <param name="data">The <see cref="XCData"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitCData([NotNull] XCData data) => new XCData(data.Value);

        /// <summary>
        /// Visits an <see cref="XComment "/>.
        /// </summary>
        /// <param name="comment">The <see cref="XComment "/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitComment([NotNull] XComment comment) => new XComment(comment.Value);

        /// <summary>
        /// Visits the <see cref="XContainer"/>.
        /// This method can dispatch to <see cref="VisitElement"/> and <see cref="VisitDocument"/>.
        /// </summary>
        /// <param name="container">The <see cref="XContainer"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        /// <exception cref="VisitorException">No dispatch was found for the derived type.</exception>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitContainer([NotNull] XContainer container)
        {
            switch (container)
            {
                case XElement e:
                    return VisitElement(e);

                case XDocument d:
                    return VisitDocument(d);

                default:
                    throw new VisitorException(container.GetType());
            }
        }

        /// <summary>
        /// Visits the <see cref="XDeclaration"/>.
        /// </summary>
        /// <param name="declaration">The <see cref="XDeclaration"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XDeclaration VisitDeclaration([NotNull] XDeclaration declaration)
            => new XDeclaration(declaration.Version, declaration.Encoding, declaration.Standalone);

        /// <summary>
        /// Visits the <see cref="XDocument"/>.
        /// </summary>
        /// <param name="document">The <see cref="XDocument"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDocument([NotNull] XDocument document)
            => new XDocument(VisitDeclaration(document.Declaration), Visit(document.Nodes()));

        /// <summary>
        /// Visits the <see cref="XDocumentType"/>.
        /// </summary>
        /// <param name="type">The <see cref="XDocumentType"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDocumentType([NotNull] XDocumentType type)
            => new XDocumentType(type.Name, type.PublicId, type.SystemId, type.InternalSubset);

        /// <summary>
        /// Visits an <see cref="XElement"/>.
        /// </summary>
        /// <param name="element">The <see cref="XElement"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XElement"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitElement([NotNull] XElement element)
            => new XElement(VisitName(element.Name), Visit(element.Attributes()), Visit(element.Nodes()));

        /// <summary>
        /// Visits an <see cref="XName"/>.
        /// </summary>
        /// <param name="name">The name to visit.</param>
        /// <returns>
        /// A visited <see cref="XName"/>.
        /// </returns>
        [Pure]
        [NotNull]
        protected virtual XName VisitName([NotNull] XName name) => VisitNamespace(name.Namespace) + name.LocalName;

        /// <summary>
        /// Visits an <see cref="XNamespace"/>.
        /// </summary>
        /// <param name="xNamespace">The namespace to visit.</param>
        /// <returns>
        /// A visited <see cref="XName"/>.
        /// </returns>
        [Pure]
        [NotNull]
        protected virtual XNamespace VisitNamespace([NotNull] XNamespace xNamespace) => xNamespace;

        /// <summary>
        /// Visits the <see cref="XNode"/>.
        /// This method can dispatch to <see cref="VisitComment"/>, <see cref="VisitContainer"/>,
        /// <see cref="VisitDocumentType"/>, <see cref="VisitProcessingInstruction"/>, and <see cref="VisitText"/>.
        /// </summary>
        /// <param name="node">The <see cref="XNode"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        /// <exception cref="VisitorException">No dispatch was found for the derived type.</exception>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitNode([NotNull] XNode node)
        {
            switch (node)
            {
                case XComment c:
                    return VisitComment(c);

                case XContainer c:
                    return VisitContainer(c);

                case XDocumentType d:
                    return VisitDocumentType(d);

                case XProcessingInstruction p:
                    return VisitProcessingInstruction(p);

                case XText t:
                    return VisitText(t);

                default:
                    throw new VisitorException(node.GetType());
            }
        }

        /// <summary>
        /// Visits an <see cref="XObject"/>.
        /// This method can dispatch to <see cref="VisitAttribute"/> and <see cref="VisitNode"/>.
        /// </summary>
        /// <param name="xObject">The <see cref="XObject"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        /// <exception cref="VisitorException">No dispatch was found for the derived type.</exception>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitObject([NotNull] XObject xObject)
        {
            switch (xObject)
            {
                case XAttribute a:
                    return VisitAttribute(a);

                case XNode n:
                    return VisitNode(n);

                default:
                    throw new VisitorException(xObject.GetType());
            }
        }

        /// <summary>
        /// Visits an <see cref="XProcessingInstruction"/>.
        /// </summary>
        /// <param name="instruction">The <see cref="XProcessingInstruction"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitProcessingInstruction([NotNull] XProcessingInstruction instruction)
            => new XProcessingInstruction(instruction.Target, instruction.Data);

        /// <summary>
        /// Visits the <see cref="string"/> as <see cref="XText"/>.
        /// </summary>
        /// <param name="text">The <see cref="string"/> to visit.</param>
        /// <returns>
        /// A visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitString([CanBeNull] string text) => text is null ? null : Visit(new XText(text));

        /// <summary>
        /// Visits an <see cref="XText"/>.
        /// This method can dispatch to <see cref="VisitCData"/>.
        /// </summary>
        /// <param name="text">The <see cref="XText"/> to visit.</param>
        /// <returns>
        /// The visited <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitText([NotNull] XText text)
        {
            switch (text)
            {
                case XCData c:
                    return VisitCData(c);

                default:
                    return new XText(text.Value);
            }
        }

        #endregion

        #region Utilities

        /// <summary>
        /// Constructs a liftable div element.
        /// </summary>
        /// <param name="element">The <see cref="XElement"/> from which nodes can be lifted.</param>
        /// <returns>
        /// An <see cref="XObject"/> representing the lift operation.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected XObject MakeLiftable([CanBeNull] XElement element)
            => element is null
                   ? null
                   : new XElement("div",
                       new XAttribute(Liftable, $"from-{element.Name.LocalName}"),
                       Visit(element.Nodes()));

        /// <summary>
        /// Yields the <see cref="XObject"/> or the children of the <see cref="XObject"/> if the "data-liftable" attribute is present.
        /// </summary>
        /// <param name="xObject">The <see cref="XObject"/> to visit.</param>
        /// <returns>
        /// The <see cref="XObject"/> or the children of the <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [NotNull]
        [LinqTunnel]
        [ItemNotNull]
        [CollectionAccess(CollectionAccessType.Read)]
        protected static IEnumerable<XObject> Lift([CanBeNull] XObject xObject)
        {
            if (xObject is null)
                yield break;

            if (!(xObject is XElement e))
            {
                yield return xObject;

                yield break;
            }

            if (e.Attribute(Liftable) is null)
            {
                yield return new XElement(e.Name, e.Attributes(), Lift(e.Nodes()));

                yield break;
            }

            foreach (XObject item in Lift(e.Nodes()))
            {
                yield return item;
            }
        }

        /// <summary>
        /// Yields the <see cref="XObject"/> or the children of the <see cref="XObject"/> if the "data-liftable" attribute is present.
        /// </summary>
        /// <param name="source">The collection of <see cref="XObject"/> to visit.</param>
        /// <returns>
        /// The <see cref="XObject"/> or the children of the <see cref="XObject"/>.
        /// </returns>
        [Pure]
        [NotNull]
        [LinqTunnel]
        [ItemNotNull]
        [CollectionAccess(CollectionAccessType.Read)]
        protected static IEnumerable<XObject> Lift([CanBeNull] [ItemCanBeNull] IEnumerable<XObject> source)
        {
            if (source is null)
                yield break;

            foreach (XObject item in source)
            {
                foreach (XObject lifted in Lift(item))
                {
                    yield return lifted;
                }
            }
        }

        /// <summary>
        /// Lifts the first node or null of an <see cref="XContainer"/> when it is empty or has a single node.
        /// </summary>
        /// <param name="container">The <see cref="XContainer"/> to visit and lift.</param>
        /// <returns>
        /// The inner <see cref="XNode"/> or the <see cref="XContainer"/>.
        /// </returns>
        [Pure]
        [CanBeNull]
        protected static XObject LiftSingleton([NotNull] XContainer container)
            => container.Nodes().Count() <= 1 ? container.FirstNode : container;

        /// <summary>
        /// Yields nodes following <paramref name="current"/> while the <paramref name="predicate"/> is true.
        /// </summary>
        /// <param name="current">The current node.</param>
        /// <param name="predicate">The condition to test subsequent nodes.</param>
        /// <returns>
        /// Nodes following <paramref name="current"/> until a node is found that fails the <paramref name="predicate"/>.
        /// </returns>
        [Pure]
        [NotNull]
        [LinqTunnel]
        [ItemNotNull]
        [CollectionAccess(CollectionAccessType.Read)]
        protected static IEnumerable<XNode> NextWhile([NotNull] XNode current, [NotNull] Func<XNode, bool> predicate)
        {
            for (XNode n = current.NextNode; n != null; n = n.NextNode)
            {
                if (!predicate(n))
                    yield break;

                yield return n;
            }
        }

        #endregion

        #region Exceptions

        /// <summary>
        /// The exception that is thrown when the visitor encounters a derived type that is not supported.
        /// </summary>
        /// <inheritdoc />
        protected class VisitorException : NotSupportedException
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="VisitorException" /> class for a specified type.
            /// </summary>
            /// <param name="type">The type that could not be dispatched.</param>
            /// <param name="callerName">The method that could not dispatch the type.</param>
            /// <inheritdoc />
            public VisitorException(Type type, [CallerMemberName] string callerName = default)
                : base($"No dispatch was found in '${callerName}' for derived type '${type.Name}'.")
            {
            }
        }

        #endregion
    }
}