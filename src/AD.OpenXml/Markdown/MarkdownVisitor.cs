using System;
using JetBrains.Annotations;

// ReSharper disable ClassWithVirtualMembersNeverInherited.Global
namespace AD.OpenXml.Markdown
{
    /// <summary>
    /// Visits text to construct a graph of Markdown nodes.
    /// </summary>
    [PublicAPI]
    public class MarkdownVisitor
    {
        /// <summary>
        /// Visits a span and returns an appropriate <see cref="MNode"/>.
        /// </summary>
        /// <param name="span">The span to visit.</param>
        /// <returns>
        /// An <see cref="MNode"/> representing the segment.
        /// </returns>
        public MNode Visit(in ReadOnlySpan<char> span)
        {
            if (MHeading.Accept(in span))
                return new MHeading(in span);

            if (MListItem.Accept(in span))
                return new MListItem(in span);

            return new MParagraph(in span);
        }
    }
}