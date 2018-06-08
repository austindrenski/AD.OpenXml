using System;
using AD.OpenXml.Markdown;
using JetBrains.Annotations;
using Xunit;

namespace AD.OpenXml.Tests
{
    [PublicAPI]
    public class MarkdownVisitorTest
    {
        [Theory]
        [InlineData(" # ", "Heading 1")]
        [InlineData("  #  ", "Heading 1")]
        [InlineData("   #   ", "Heading 1")]
        [InlineData("# ", "Heading 1 ###### ")]
        [InlineData("# ", "Heading 1")]
        [InlineData("## ", "Heading 2")]
        [InlineData("### ", "Heading 3")]
        [InlineData("#### ", "Heading 4")]
        [InlineData("##### ", "Heading 5")]
        [InlineData("###### ", "Heading 6")]
        public void HeadingPass(string prefix, string value)
        {
            MarkdownVisitor visitor = new MarkdownVisitor();
            ReadOnlySpan<char> raw = prefix + value;
            string result = value.Trim().TrimEnd('#').TrimEnd();
            string markdown = $"{prefix.Trim()} {result}";

            MNode node = visitor.Visit(in raw);

            Assert.IsType<MHeading>(node);
            Assert.Equal(markdown, node.ToString());
            Assert.Equal(result, ((MHeading) node).Heading.ToString());
        }

        [Theory]
        [InlineData("# ")]
        [InlineData("    # Heading 1")]
        [InlineData("Heading 1")]
        [InlineData("#Heading 1")]
        [InlineData("##Heading 2")]
        public void HeadingFail(string text)
        {
            MarkdownVisitor visitor = new MarkdownVisitor();
            ReadOnlySpan<char> span = text;
            string markdown = text.Trim();

            MNode node = visitor.Visit(in span);

            Assert.IsNotType<MHeading>(node);
            Assert.Equal(markdown, node.ToString());
        }
    }
}