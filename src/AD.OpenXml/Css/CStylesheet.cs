using System;
using System.Linq;
using System.Text;
using JetBrains.Annotations;

namespace AD.OpenXml.Css
{
    public class CStylesheet
    {
        [NotNull] private readonly string _name;
        [NotNull] [ItemNotNull] private readonly CRuleset[] _rulesets;

        public CStylesheet(ReadOnlySpan<char> name, [NotNull] [ItemCanBeNull] params CRuleset[] rulesets)
        {
            if (name.Contains("*/", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException($"{nameof(name)} cannot contain '*/': {name.ToString()}");
            if (rulesets is null)
                throw new ArgumentNullException(nameof(rulesets));

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