using System;
using JetBrains.Annotations;

namespace AD.OpenXml.Css
{
    /// <summary>
    /// Represents a selector for a CSS ruleset.
    /// </summary>
    /// <remarks>
    /// See: https://developer.mozilla.org/en-US/docs/Web/CSS/Reference
    /// </remarks>
    [PublicAPI]
    public class CSelector
    {
        /// <summary>
        /// The CSS selector.
        /// </summary>
        [NotNull] readonly string _selector;

        /// <summary>
        /// Initializes a <see cref="CSelector"/> from the selector.
        /// </summary>
        /// <param name="selector">The CSS selector.</param>
        /// <exception cref="ArgumentException">Invalid CSS selector.</exception>
        public CSelector(in ReadOnlySpan<char> selector)
        {
            if (selector.IsEmpty)
                throw new ArgumentException($"Invalid CSS selector: {selector.ToString()}");

            _selector = selector.ToString();
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString() => _selector;
    }
}