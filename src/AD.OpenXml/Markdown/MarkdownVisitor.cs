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
            if (MHeading.Accept(span))
                return new MHeading(span);

            if (MBulletListItem.Accept(span))
                return new MBulletListItem(span);

            return new MParagraph(span);
        }
    }
}