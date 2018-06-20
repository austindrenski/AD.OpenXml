using System;
using System.Linq;
using System.Text;
using JetBrains.Annotations;

namespace AD.OpenXml.Css
{
    /// <summary>
    /// Represents a CSS stylesheet.
    /// </summary>
    /// <remarks>
    /// See: https://developer.mozilla.org/en-US/docs/Web/CSS/Reference
    /// </remarks>
    [PublicAPI]
    public class CStylesheet
    {
        /// <summary>
        /// The name of the stylesheet.
        /// </summary>
        [NotNull] private readonly string _name;

        /// <summary>
        /// The rulesets in the stylesheet.
        /// </summary>
        [NotNull] [ItemNotNull] private readonly CRuleset[] _rulesets;

        /// <summary>
        /// Initializes a <see cref="CDeclaration"/> from the property and value.
        /// </summary>
        /// <param name="name">The CSS property name.</param>
        /// <param name="rulesets">The CSS property value.</param>
        /// <exception cref="ArgumentNullException" />
        /// <exception cref="ArgumentException">Invalid stylesheet name.</exception>
        public CStylesheet(ReadOnlySpan<char> name, [NotNull] [ItemCanBeNull] params CRuleset[] rulesets)
        {
            if (rulesets is null)
                throw new ArgumentNullException(nameof(rulesets));
            if (name.IsEmpty || name.Contains("*/", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException($"{nameof(name)} cannot be empty or contain '*/': {name.ToString()}");


            _name = name.ToString();
            _rulesets = rulesets.Where(x => x != null).ToArray();
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine($"/* {_name} */");

            for (int i = 0; i < _rulesets.Length; i++)
            {
                sb.AppendLine(_rulesets[i].ToString());
            }

            return sb.ToString();
        }
    }
}