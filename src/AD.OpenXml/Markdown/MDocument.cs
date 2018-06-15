using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using JetBrains.Annotations;

namespace AD.OpenXml.Markdown
{
    /// <inheritdoc />
    /// <summary>
    /// Represents a Markdown document.
    /// </summary>
    /// <remarks>
    /// See: http://spec.commonmark.org
    /// </remarks>
    [PublicAPI]
    public class MDocument : MNode
    {
        private readonly List<MNode> _nodes = new List<MNode>();

        /// <summary>
        ///
        /// </summary>
        /// <param name="node"></param>
        public void Append([NotNull] MNode node) => _nodes.Add(node);

        /// <inheritdoc />
        [Pure]
        public override XNode ToHtml()
            => (XNode) new HtmlVisitor().Visit(ToOpenXml()) ?? throw new ArgumentException();

        /// <inheritdoc />
        [Pure]
        public override XNode ToOpenXml()
            => new XElement(W + "document",
                new XElement(W + "body",
                    _nodes.Select(x => x.ToOpenXml())));
    }
}