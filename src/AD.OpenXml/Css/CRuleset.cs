using System;
using System.Text;
using JetBrains.Annotations;

namespace AD.OpenXml.Css
{
    public class CRuleset
    {
        [NotNull] private readonly CSelector _selector;
        [NotNull] [ItemCanBeNull] private readonly CDeclaration[] _declarations;

        public CRuleset([NotNull] CSelector selector, [NotNull] [ItemCanBeNull] params CDeclaration[] declarations)
        {
            if (declarations is null)
                throw new ArgumentNullException(nameof(declarations));

            _selector = selector;
            _declarations = declarations;
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
                if (_declarations[i] is CDeclaration declaration)
                    sb.AppendLine($"    {declaration}");
            }

            sb.AppendLine("}");

            return sb.ToString();
        }
    }
}