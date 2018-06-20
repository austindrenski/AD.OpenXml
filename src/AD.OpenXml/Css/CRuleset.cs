using System;
using System.Linq;
using System.Text;
using JetBrains.Annotations;

namespace AD.OpenXml.Css
{
    /// <summary>
    /// Represents a CSS ruleset.
    /// </summary>
    /// <remarks>
    /// See: https://developer.mozilla.org/en-US/docs/Web/CSS/Reference
    /// </remarks>
    [PublicAPI]
    public class CRuleset
    {
        /// <summary>
        /// The CSS selector.
        /// </summary>
        [NotNull] private readonly CSelector _selector;

        /// <summary>
        /// The rule declarations.
        /// </summary>
        [NotNull] [ItemNotNull] private readonly CDeclaration[] _declarations;

        /// <summary>
        /// Initializes a <see cref="CRuleset"/> from the selector and declarations.
        /// </summary>
        /// <param name="selector">The CSS selector.</param>
        /// <param name="declarations">The CSS rule declarations.</param>
        /// <exception cref="ArgumentNullException" />
        public CRuleset([NotNull] CSelector selector, [NotNull] [ItemCanBeNull] params CDeclaration[] declarations)
        {
            if (selector is null)
                throw new ArgumentNullException(nameof(selector));
            if (declarations is null)
                throw new ArgumentNullException(nameof(declarations));

            _selector = selector;
            _declarations = declarations.Where(x => x != null).ToArray();
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(_selector);

            sb.AppendLine(" {");

            for (int i = 0; i < _declarations.Length; i++)
            {
                sb.Append("    ");
                sb.AppendLine(_declarations[i].ToString());
            }

            sb.AppendLine("}");

            return sb.ToString();
        }
    }
}