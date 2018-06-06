using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;

namespace AD.OpenXml.Markdown
{
    /// <summary>
    /// Visits text to construct a graph of Markdown nodes.
    /// </summary>
    [PublicAPI]
    public class MarkdownVisitor
    {
        /// <summary>
        /// Visits a <see cref="StringSegment"/> and returns an appropriate <see cref="MNode"/>.
        /// </summary>
        /// <param name="segment">
        /// The segment to visit.
        /// </param>
        /// <returns>
        /// An <see cref="MNode"/> representing the segment.
        /// </returns>
        public MNode Visit(in StringSegment segment)
        {
            if (MHeading.Accept(in segment))
                return new MHeading(in segment);

            return new MParagraph(in segment);
        }
    }
}