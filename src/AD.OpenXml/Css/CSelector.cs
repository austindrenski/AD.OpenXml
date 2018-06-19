using System;
using JetBrains.Annotations;

namespace AD.OpenXml.Css
{
    public class CSelector
    {
        [NotNull] private readonly string _selector;

        public CSelector([NotNull] string selector)
        {
            if (selector is null)
                throw new ArgumentNullException(nameof(selector));

            _selector = selector;
        }

        /// <inheritdoc />
        [Pure]
        public override string ToString() => _selector;
    }
}