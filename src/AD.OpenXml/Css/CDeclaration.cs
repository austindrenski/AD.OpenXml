using System;
using JetBrains.Annotations;

namespace AD.OpenXml.Css
{
    /// <summary>
    /// Represents a declaration in a CSS rule.
    /// </summary>
    /// <remarks>
    /// See: https://developer.mozilla.org/en-US/docs/Web/CSS/Reference
    /// </remarks>
    [PublicAPI]
    public class CDeclaration
    {
        /// <summary>
        /// The CSS property name.
        /// </summary>
        [NotNull] readonly string _property;

        /// <summary>
        /// The CSS property value.
        /// </summary>
        [NotNull] readonly string _value;

        /// <summary>
        /// Initializes a <see cref="CDeclaration"/> from the property and value.
        /// </summary>
        /// <param name="property">The CSS property name.</param>
        /// <param name="value">The CSS property value.</param>
        /// <exception cref="ArgumentException">Invalid CSS property name.</exception>
        /// <exception cref="ArgumentException">Invalid CSS property value.</exception>
        public CDeclaration(in ReadOnlySpan<char> property, in ReadOnlySpan<char> value)
        {
            ReadOnlySpan<char> p = property.Trim();
            ReadOnlySpan<char> v = value.Trim();

            if (!ValidateProperty(in p))
                throw new ArgumentException($"Invalid CSS property name: {p.ToString()}");

            if (v.IsEmpty)
                throw new ArgumentException($"Invalid CSS property value: {v.ToString()}");

            _property = p.ToString();
            _value = v.ToString();
        }

        /// <inheritdoc />
        [Pure]
        [NotNull]
        public override string ToString() => $"{_property}: {_value};";

        /// <summary>
        /// Validates CSS property names.
        /// </summary>
        /// <param name="span">The property name to validate.</param>
        /// <returns>
        /// True if the property name is valid; otherwise false.
        /// </returns>
        /// <remarks>
        /// See: https://drafts.csswg.org/css-syntax/#syntax-description
        /// </remarks>
        [Pure]
        static bool ValidateProperty(in ReadOnlySpan<char> span)
        {
            if (span.IsEmpty)
                return false;

            if (char.IsDigit(span[0]))
                return false;

            if (span[0] == '-' && (span.Length == 1 || char.IsDigit(span[2])))
                return false;

            for (int i = 0; i < span.Length; i++)
            {
                if (char.IsLetterOrDigit(span[i]))
                    continue;

                if (span[i] == '-' || span[i] == '_')
                    continue;

                return false;
            }

            return true;
        }
    }
}