using System;
using System.Collections.Generic;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

// ReSharper disable ClassWithVirtualMembersNeverInherited.Global

namespace AD.OpenXml
{
    /// <inheritdoc />
    /// <summary>
    /// Represents a <see cref="MathMLVisitor"/> that can transform an OpenXML math node into a well-formed MathML node.
    /// </summary>
    public class MathMLVisitor : XmlVisitor
    {
        /// <summary>
        /// Return an element when handling the default dispatch case if true; otherwise, false.
        /// </summary>
        private readonly bool _returnOnDefault;

        // TODO: move into AD.Xml
        /// <summary>
        /// Represents the 'm:' prefix seen in the markup for math elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

        /// <summary>
        /// Represents the 'w:' prefix seen in raw OpenXML documents.
        /// </summary>
        [NotNull] protected static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        /// <summary>
        /// The mapping of <see cref="XName"/> to visit method used by <see cref="VisitElement"/>.
        /// </summary>
        protected readonly IDictionary<XName, Func<XElement, XObject>> VisitLookup;

        /// <summary>
        /// The supported operator characters.
        /// </summary>
        protected virtual ISet<string> SupportedOperators { get; } =
            new HashSet<string>
            {
                "(",
                ")",
                ",",
                "+",
                "-",
                "=",
                "!",
                "exp",
                "log"
            };

        /// <summary>
        /// Initializes an <see cref="MathMLVisitor"/>.
        /// </summary>
        /// <param name="returnOnDefault">
        /// True if an element should be returned when handling the default dispatch case.
        /// </param>
        protected MathMLVisitor(bool returnOnDefault)
        {
            _returnOnDefault = returnOnDefault;

            VisitLookup =
                new Dictionary<XName, Func<XElement, XObject>>
                {
                    // @formatter:off
                    [M + "d"]         = VisitDelimiter,
                    [M + "e"]         = VisitBase,
                    [M + "oMath"]     = VisitMath,
                    [M + "oMathPara"] = VisitMathParagraph,
                    [M + "r"]         = VisitRun,
                    [M + "sSub"]      = VisitSubscript,
                    [M + "sSup"]      = VisitSuperscript,
                    // @formatter:on
                };
        }

        /// <summary>
        /// Returns a new <see cref="MathMLVisitor"/>.
        /// </summary>
        /// <param name="returnOnDefault">
        /// True if an element should be returned when handling the default dispatch case.
        /// </param>
        /// <returns>
        /// A <see cref="MathMLVisitor"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [NotNull]
        public static MathMLVisitor Create(bool returnOnDefault = false)
        {
            return new MathMLVisitor(returnOnDefault);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="baseItem">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitBase([NotNull] XElement baseItem)
        {
            if (baseItem is null)
            {
                throw new ArgumentNullException(nameof(baseItem));
            }

            return LiftableHelper(baseItem);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="delimiter">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDelimiter([NotNull] XElement delimiter)
        {
            if (delimiter is null)
            {
                throw new ArgumentNullException(nameof(delimiter));
            }

            return
                new XElement("mrow",
                    new XElement("mo",
                        new XAttribute("fence", "true"),
                        new XText("(")),
                    Visit(delimiter.Nodes()),
                    new XElement("mo",
                        new XAttribute("fence", "true"),
                        new XText(")")));
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

            return
                new XElement("math",
                    Visit(math.Nodes()));
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="mathParagraph">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitMathParagraph([NotNull] XElement mathParagraph)
        {
            if (mathParagraph is null)
            {
                throw new ArgumentNullException(nameof(mathParagraph));
            }

            return LiftableHelper(mathParagraph);
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
        protected virtual XObject VisitRun([NotNull] XElement run)
        {
            if (run is null)
            {
                throw new ArgumentNullException(nameof(run));
            }

            string value = run.Value.Trim();

            return
                SupportedOperators.Contains(value)
                    ? new XElement("mo", value)
                    : run.Parent.Name == M + "oMath"
                        ? new XElement("mi", value)
                        : Visit(value);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="subscript">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSubscript([NotNull] XElement subscript)
        {
            if (subscript is null)
            {
                throw new ArgumentNullException(nameof(subscript));
            }

            return
                new XElement("msub",
                    LiftableHelper(subscript.Element(M + "e")),
                    LiftableHelper(subscript.Element(M + "sub")));
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="superscript">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSuperscript([NotNull] XElement superscript)
        {
            if (superscript is null)
            {
                throw new ArgumentNullException(nameof(superscript));
            }

            return
                new XElement("msup",
                    LiftableHelper(superscript.Element(M + "e")),
                    LiftableHelper(superscript.Element(M + "sup")));
        }

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitText(XText text)
        {
            if (text is null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            return
                double.TryParse(text.Value, out double _)
                    ? new XElement("mn", text)
                    : new XElement("mi", text);
        }
    }
}