using System;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Markdown
{
    /// <inheritdoc cref="MListItem"/>
    /// <inheritdoc cref="IEquatable{T}"/>
    /// <summary>
    /// Represents a Markdown list item node.
    /// </summary>
    /// <remarks>
    /// See: https://spec.commonmark.org/0.28/#lists
    /// </remarks>
    [PublicAPI]
    public sealed class MBulletListItem : MListItem
    {
        /// <inheritdoc />
        public override int Level { get; }

        /// <inheritdoc />
        public override char Marker { get; }

        /// <inheritdoc />
        public override MText Item { get; }

        /// <summary>
        /// Constructs an <see cref="MBulletListItem"/>.
        /// </summary>
        /// <param name="text">The raw text of the item.</param>
        public MBulletListItem(in ReadOnlySpan<char> text)
        {
            if (!Accept(text))
                throw new ArgumentException($"Lists must begin with '-', '*', or '+': '{text.ToString()}'");

            Marker = text[text.IndexOfAny('-', '*', '+')];
            Level = text.IndexOf(Marker) / 2;
            Item = Normalize(text);
        }

        /// <summary>
        /// Checks if the segment is a well-formed Markdown heading.
        /// </summary>
        /// <param name="span">The span to test.</param>
        /// <returns>
        /// True if the segment is a well-formed Markdown heading; otherwise false.
        /// </returns>
        [Pure]
        public new static bool Accept(in ReadOnlySpan<char> span)
        {
            ReadOnlySpan<char> trimmed = span.Trim();

            if (trimmed.Length < 2)
                return false;

            if (trimmed[0] != '-' &&
                trimmed[0] != '*' &&
                trimmed[0] != '+')
                return false;

            // If at least 1 space does not follow the marker, it is not a list.
            if (trimmed[1] != ' ')
                return false;

            // If 5 (or more) spaces follow the marker, it is not a list.
            return !trimmed.Slice(1).StartsWith("     ");
        }

        /// <summary>
        /// Normalizes the segment by trimming (in order):
        ///   1) ' ' from the start and end;
        ///   2) one of the following from the start:
        ///        - '- '
        ///        - '#) '
        ///        - '#. '
        ///   3) ' ' from the start;
        ///   4) normalizing inner whitespace.
        /// </summary>
        /// <param name="span">The span to normalize.</param>
        /// <returns>
        /// The normalized segment.
        /// </returns>
        [Pure]
        private static ReadOnlySpan<char> Normalize(in ReadOnlySpan<char> span)
            => span.Slice(span.IndexOfAny('-', '*', '+') + 1).Trim();

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString() => $"{new string(' ', 2 * Level)}{Marker} {Item}";

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml() => new XElement("li", Item.ToHtml());

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml()
            => new XElement(W + "p",
                new XElement(W + "pPr",
                    new XElement(W + "pStyle",
                        new XAttribute(W + "val", "ListParagraph"))),
                new XElement(W + "r", Item.ToOpenXml()));
    }
}