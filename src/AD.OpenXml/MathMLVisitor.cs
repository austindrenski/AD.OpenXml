using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

namespace AD.OpenXml
{
    // ReSharper disable ClassWithVirtualMembersNeverInherited.Global
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

        /// <summary>
        /// Represents the 'm:' prefix seen in the markup for math elements.
        /// </summary>
        [NotNull] protected static readonly XNamespace M = XNamespaces.OpenXmlMath;

        /// <summary>
        /// The mapping of OMML accent characters to MathML accent characters.
        /// </summary>
        protected readonly IDictionary<string, string> AccentLookup =
            new Dictionary<string, string>
            {
                ["\u0305"] = "\u02C9"
            };

        /// <summary>
        /// The supported operator characters.
        /// </summary>
        protected readonly HashSet<string> SupportedOperators =
            new HashSet<string>
            {
                "(",
                ")",
                "[",
                "]",
                "{",
                "{",
                ",",
                "+",
                "-",
                "*",
                "/",
                "=",
                "!",
                "exp",
                "log",
                "sin",
                "cos",
                "acos"
            };

        /// <summary>
        /// The ignored nodes.
        /// </summary>
        protected readonly HashSet<XName> IgnoredNodes =
            new HashSet<XName>
            {
                M + "accPr",
                M + "ctrlPr",
                M + "dPr",
                M + "eqArrPr",
                M + "fPr",
                M + "mPr",
                M + "naryPr",
                M + "rPr",
                M + "sSubPr",
                M + "sSubSupPr",
                M + "sSupPr"
            };

        /// <summary>
        /// Initializes an <see cref="MathMLVisitor"/>.
        /// </summary>
        /// <param name="returnOnDefault">True to return an element from the default dispatch case.</param>
        protected MathMLVisitor(bool returnOnDefault)
            => _returnOnDefault = returnOnDefault;

        /// <summary>
        /// Returns a new <see cref="MathMLVisitor"/>.
        /// </summary>
        /// <param name="returnOnDefault">True to return an element from the default dispatch case.</param>
        /// <returns>
        /// A <see cref="MathMLVisitor"/>.
        /// </returns>
        [Pure]
        [NotNull]
        public static MathMLVisitor Create(bool returnOnDefault = false)
            => new MathMLVisitor(returnOnDefault);

        /// <summary>
        ///
        /// </summary>
        /// <param name="accent"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitAccent([NotNull] XElement accent)
            => accent.Element(M + "accPr")?.Element(M + "chr")?.Attribute(M + "val")?.Value is string value
                   ? new XElement("mover",
                       new XAttribute("accent", true),
                       Visit(accent.Nodes()),
                       new XElement("mo", AccentLookup.TryGetValue(value, out string mapped) ? mapped : value))
                   : null;

        /// <summary>
        ///
        /// </summary>
        /// <param name="array"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitArray([NotNull] XElement array)
            => new XElement("mtable",
                Visit(array.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="bar"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitBar([NotNull] XElement bar)
            => new XElement("mover",
                new XAttribute("accent", true),
                Visit(bar.Nodes()),
                new XElement("mo", '\u02C9'));

        /// <summary>
        ///
        /// </summary>
        /// <param name="baseItem"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitBase([NotNull] XElement baseItem)
            => new XElement("mrow", Visit(baseItem.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="delimiter"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDelimiter([NotNull] XElement delimiter)
            => new XElement("mrow",
                new XElement("mo",
                    new XAttribute("fence", true),
                    new XText(delimiter.Element(M + "dPr")?.Element(M + "begChr")?.Attribute(M + "val")?.Value ?? "(")),
                new XElement("mrow",
                    Visit(delimiter.Nodes())),
                new XElement("mo",
                    new XAttribute("fence", true),
                    new XText(delimiter.Element(M + "dPr")?.Element(M + "endChr")?.Attribute(M + "val")?.Value ?? ")")));

        /// <summary>
        ///
        /// </summary>
        /// <param name="degree"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDegree([NotNull] XElement degree)
            => new XElement("mrow", Visit(degree.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="denominator"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitDenominator([NotNull] XElement denominator)
            => new XElement("mrow", Visit(denominator.Nodes()));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitElement(XElement element)
        {
            if (element.Name.Namespace == XNamespace.None)
                return element;

            if (element.Name.Namespace != M)
                return null;

            switch (element.Name.LocalName)
            {
                case "acc":
                    return VisitAccent(element);

                case "bar":
                    return VisitBar(element);

                case "d":
                    return VisitDelimiter(element);

                case "deg":
                    return VisitDegree(element);

                case "den":
                    return VisitDenominator(element);

                case "e":
                    return VisitBase(element);

                case "eqArr":
                    return VisitArray(element);

                case "f":
                    return VisitFraction(element);

                case "fName":
                    return VisitFunctionName(element);

                case "func":
                    return VisitFunction(element);

                case "lim":
                    return VisitLimit(element);

                case "m":
                    return VisitMatrix(element);

                case "mr":
                    return VisitMatrixRow(element);

                case "num":
                    return VisitNumerator(element);

                case "nary":
                    return VisitNary(element);

                case "oMath":
                    return VisitMath(element);

                case "oMathPara":
                    return VisitMathParagraph(element);

                case "r":
                    return VisitRun(element);

                case "rad":
                    return VisitRoot(element);

                case "sSub":
                    return VisitSubscript(element);

                case "sSubSup":
                    return VisitSubscriptSuperscript(element);

                case "sSup":
                    return VisitSuperscript(element);

                case "sub":
                    return VisitSubscriptItem(element);

                case "sup":
                    return VisitSuperscriptItem(element);

                case "t":
                    return VisitText(new XText(element.Value));

                default:
                    return
                        IgnoredNodes.Contains(element.Name) || !_returnOnDefault
                            ? null
                            : base.VisitElement(element);
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="function"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFunction([NotNull] XElement function)
            => new XElement("mrow",
                Visit(function.Element(M + "fName")),
                new XElement("mo", "&ApplyFunction;"),
                Visit(function.Element(M + "e")));

        /// <summary>
        ///
        /// </summary>
        /// <param name="functionName"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFunctionName([NotNull] XElement functionName)
        {
            if (functionName.Element(M + "limLow") is XElement lowerLimit)
            {
                return
                    new XElement("munder",
                        Visit(lowerLimit.Element(M + "e")),
                        Visit(lowerLimit.Element(M + "lim")));
            }

            if (functionName.Element(M + "limUpp") is XElement upperLimit)
            {
                return
                    new XElement("mover",
                        Visit(upperLimit.Element(M + "e")),
                        Visit(upperLimit.Element(M + "lim")));
            }

            return new XElement("mi", functionName.Value);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="fraction"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitFraction([NotNull] XElement fraction)
            => new XElement("mfrac", Visit(fraction.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="limit"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitLimit([NotNull] XElement limit)
            => MakeLiftable(limit);

        /// <summary>
        ///
        /// </summary>
        /// <param name="math"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitMath([NotNull] XElement math)
            => new XElement("math", Visit(math.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="mathParagraph"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitMathParagraph([NotNull] XElement mathParagraph)
            => MakeLiftable(mathParagraph);

        /// <summary>
        ///
        /// </summary>
        /// <param name="matrix"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitMatrix([NotNull] XElement matrix)
            => new XElement("mtable", Visit(matrix.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="row"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitMatrixRow([NotNull] XElement row)
            => new XElement("mtr",
                Visit(row.Nodes()).Select(x => new XElement("mtd", x)));

        /// <summary>
        ///
        /// </summary>
        /// <param name="nary"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitNary([NotNull] XElement nary)
            => new XElement("mrow",
                new XElement("munderover",
                    new XElement("mo",
                        nary.Element(M + "naryPr")?
                           .Element(M + "chr")?
                           .Attribute(M + "val")?
                           .Value ?? "&Integral;"),
                    Visit(nary.Element(M + "sub")),
                    Visit(nary.Element(M + "sup"))),
                Visit(nary.Element(M + "e")));

        /// <summary>
        ///
        /// </summary>
        /// <param name="numerator"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitNumerator([NotNull] XElement numerator)
            => new XElement("mrow", Visit(numerator.Nodes()));

        /// <summary>
        ///
        /// </summary>
        /// <param name="root"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitRoot([NotNull] XElement root)
            => root.Element(M + "deg") is XElement degree && degree.HasElements
                   ? new XElement("mroot",
                       Visit(root.Element(M + "e")),
                       Visit(degree))
                   : new XElement("msqrt",
                       Visit(root.Element(M + "e")));

        /// <summary>
        ///
        /// </summary>
        /// <param name="run"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitRun([NotNull] XElement run)
            => MakeLiftable(run);

        /// <summary>
        ///
        /// </summary>
        /// <param name="subscript"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSubscript([NotNull] XElement subscript)
            => new XElement("msub",
                new XElement("mrow",
                    MakeLiftable(subscript.Element(M + "e"))),
                new XElement("mrow",
                    MakeLiftable(subscript.Element(M + "sub"))));

        /// <summary>
        ///
        /// </summary>
        /// <param name="subscriptItem"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSubscriptItem([NotNull] XElement subscriptItem)
            => new XElement("mrow", MakeLiftable(subscriptItem));

        /// <summary>
        ///
        /// </summary>
        /// <param name="subscriptSuperscript"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSubscriptSuperscript([NotNull] XElement subscriptSuperscript)
            => new XElement("msubsup",
                new XElement("mrow",
                    MakeLiftable(subscriptSuperscript.Element(M + "e"))),
                new XElement("mrow",
                    MakeLiftable(subscriptSuperscript.Element(M + "sub"))),
                new XElement("mrow",
                    MakeLiftable(subscriptSuperscript.Element(M + "sup"))));

        /// <summary>
        ///
        /// </summary>
        /// <param name="superscript"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSuperscript([NotNull] XElement superscript)
            => new XElement("msup",
                new XElement("mrow",
                    MakeLiftable(superscript.Element(M + "e"))),
                new XElement("mrow",
                    MakeLiftable(superscript.Element(M + "sup"))));

        /// <summary>
        ///
        /// </summary>
        /// <param name="superscriptItem"></param>
        /// <returns>
        ///
        /// </returns>
        [Pure]
        [CanBeNull]
        protected virtual XObject VisitSuperscriptItem([NotNull] XElement superscriptItem)
            => new XElement("mrow", MakeLiftable(superscriptItem));

        /// <inheritdoc />
        [Pure]
        protected override XObject VisitText(XText text)
        {
            string value = text.Value.Trim();

            if (string.IsNullOrWhiteSpace(value))
                return null;

            if (SupportedOperators.Contains(value))
                return new XElement("mo", value);

            return
                double.TryParse(text.Value, out double d)
                    ? new XElement("mn", d)
                    : new XElement("mi", text.Value);
        }
    }
}