using AD.OpenXml.Markdown;
using JetBrains.Annotations;
using Microsoft.Extensions.Primitives;
using Xunit;

namespace AD.OpenXml.Tests
{
    [PublicAPI]
    public class MarkdownVisitorTest
    {
        [Theory]
        [InlineData(" #  ", "Heading 1")]
        [InlineData("#  ", "Heading 1")]
        [InlineData("# ", "Heading 1")]
        [InlineData("## ", "Heading 2")]
        [InlineData("### ", "Heading 3")]
        [InlineData("#### ", "Heading 4")]
        [InlineData("##### ", "Heading 5")]
        [InlineData("###### ", "Heading 6")]
        public void HeadingPass(string prefix, string value)
        {
            MarkdownVisitor visitor = new MarkdownVisitor();
            StringSegment raw = prefix + value;
            string markdown = $"{prefix.Trim()} {value.Trim()}";

            MNode node = visitor.Visit(in raw);

            Assert.IsType<MHeading>(node);

            MHeading heading = (MHeading) node;

            Assert.Equal(markdown, heading.ToString());
            Assert.Equal(value, heading.Heading.ToString());
        }

        [Theory]
        [InlineData("Heading 1")]
        [InlineData("#Heading 1")]
        [InlineData("##Heading 2")]
        public void HeadingFail(string text)
        {
            MarkdownVisitor visitor = new MarkdownVisitor();
            StringSegment segment = text;

            MNode node = visitor.Visit(in segment);

            Assert.IsNotType<MHeading>(node);
        }
    }
}