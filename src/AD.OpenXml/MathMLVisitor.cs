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
        /// The mapping of OMML accent characters to MathML accent characters.
        /// </summary>
        protected readonly IDictionary<string, string> AccentLookup;

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
        /// The ignored nodes.
        /// </summary>
        protected virtual ISet<XName> IgnoredNodes { get; } =
            new HashSet<XName>
            {
                M + "accPr",
                M + "ctrlPr",
                M + "dPr",
                M + "eqArrPr",
                M + "fPr",
                M + "sSubPr",
                M + "sSubSupPr",
                M + "sSupPr"
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

            AccentLookup =
                new Dictionary<string, string>
                {
                    ["\u0305"] = "\u02C9"
                };

            VisitLookup =
                new Dictionary<XName, Func<XElement, XObject>>
                {
                    // @formatter:off
                    [M + "acc"]       = VisitAccent,
                    [M + "d"]         = VisitDelimiter,
                    [M + "den"]       = VisitDenominator,
                    [M + "e"]         = VisitBase,
                    [M + "eqArr"]     = VisitArray,
                    [M + "f"]         = VisitFraction,
                    [M + "num"]       = VisitNumerator,
                    [M + "oMath"]     = VisitMath,
                    [M + "oMathPara"] = VisitMathParagraph,
                    [M + "r"]         = VisitRun,
                    [M + "sSub"]      = VisitSubscript,
                    [M + "sSubSup"]   = VisitSubscriptSuperscript,
                    [M + "sSup"]      = VisitSuperscript
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
        public static MathMLVisitor Create(bool returnOnDefault = false) => new MathMLVisitor(returnOnDefault);

        /// <summary>
        ///
        /// </summary>
        /// <param name="array">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitArray([NotNull] XElement array) => new XElement("mtable", Visit(array.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="accent">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAccent([NotNull] XElement accent)
        {
            string value = (string) accent.Element(M + "accPr")?.Element(M + "chr")?.Attribute(M + "val");
            return
                new XElement("mover",
                    new XAttribute("accent", true),
                    Visit(accent.Nodes()),
                    new XElement("mo", AccentLookup.TryGetValue(value, out string mapped) ? mapped : value));
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
        protected virtual XObject VisitBase([NotNull] XElement baseItem) => LiftableHelper(baseItem);

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
            => new XElement("mrow",
                new XElement("mo",
                    new XAttribute("fence", "true"),
                    new XText("(")),
                Visit(delimiter.Nodes()),
                new XElement("mo",
                    new XAttribute("fence", "true"),
                    new XText(")")));

        /// <summary>
        ///
        /// </summary>
        /// <param name="denominator">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDenominator([NotNull] XElement denominator) => new XElement("mrow", Visit(denominator.Nodes()));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitElement(XElement element)
            => VisitLookup.TryGetValue(element.Name, out Func<XElement, XObject> visit)
                   ? visit(element)
                   : IgnoredNodes.Contains(element.Name)
                       ? null
                       : _returnOnDefault
                           ? base.VisitElement(element)
                           : null;

        /// <summary>
        ///
        /// </summary>
        /// <param name="fraction">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException"/>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFraction([NotNull] XElement fraction) => new XElement("mfrac", Visit(fraction.Nodes()));

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
        protected virtual XObject VisitMath([NotNull] XElement math) => new XElement("math", Visit(math.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="numerator">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitNumerator([NotNull] XElement numerator) => new XElement("mrow", Visit(numerator.Nodes()));

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
        protected virtual XObject VisitMathParagraph([NotNull] XElement mathParagraph) => LiftableHelper(mathParagraph);

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
            string value = run.Value.Trim();

            return
                SupportedOperators.Contains(value)
                    ? new XElement("mo", value)
                    : run.Parent?.Name == M + "oMath"
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
            => new XElement("msub",
                LiftableHelper(subscript.Element(M + "e")),
                LiftableHelper(subscript.Element(M + "sub")));

        /// <summary>
        ///
        /// </summary>
        /// <param name="subscriptSuperscript">
        ///
        /// </param>
        /// <returns>
        ///
        /// </returns>
        /// <exception cref="ArgumentNullException" />
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSubscriptSuperscript([NotNull] XElement subscriptSuperscript)
            => new XElement("msubsup", Visit(subscriptSuperscript.Nodes()));

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
            => new XElement("msup",
                LiftableHelper(superscript.Element(M + "e")),
                LiftableHelper(superscript.Element(M + "sup")));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitText(XText text)
            => string.IsNullOrWhiteSpace(text.Value)
                   ? null
                   : double.TryParse(text.Value, out double value)
                       ? new XElement("mn", value)
                       : new XElement("mi", text);
    }
}