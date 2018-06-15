using System.Linq;
using AD.OpenXml.Markdown;
using JetBrains.Annotations;
using Xunit;

namespace AD.OpenXml.Tests
{
    [PublicAPI]
    public class MarkdownVisitorTest
    {
        [Theory]
        [InlineData(" # Heading 1", "# Heading 1", "Heading 1", "<h1>Heading 1</h1>")]
        [InlineData("  #  Heading 1", "# Heading 1", "Heading 1", "<h1>Heading 1</h1>")]
        [InlineData("   #   Heading 1", "# Heading 1", "Heading 1", "<h1>Heading 1</h1>")]
        [InlineData("# Heading 1 ###### ", "# Heading 1", "Heading 1", "<h1>Heading 1</h1>")]
        [InlineData("# Heading 1", "# Heading 1", "Heading 1", "<h1>Heading 1</h1>")]
        [InlineData("## Heading 2", "## Heading 2", "Heading 2", "<h2>Heading 2</h2>")]
        [InlineData("### Heading 3", "### Heading 3", "Heading 3", "<h3>Heading 3</h3>")]
        [InlineData("#### Heading 4", "#### Heading 4", "Heading 4", "<h4>Heading 4</h4>")]
        [InlineData("##### Heading 5", "##### Heading 5", "Heading 5", "<h5>Heading 5</h5>")]
        [InlineData("###### Heading 6", "###### Heading 6", "Heading 6", "<h6>Heading 6</h6>")]
        public void GivenText_WhenValid_ReturnsMHeading(string markdown, string normalized, string text, string html)
        {
            MNode node = new MarkdownVisitor().Visit(markdown);

            MHeading heading = Assert.IsType<MHeading>(node);
            Assert.Equal(text, heading.Heading.ToString());
            Assert.Equal(html, heading.ToHtml().ToString());
            Assert.Equal(normalized, node.ToString());
        }

        [Theory]
        [InlineData("# ")]
        [InlineData("    # Heading 1")]
        [InlineData("Heading 1")]
        [InlineData("#Heading 1")]
        [InlineData("##Heading 2")]
        public void GivenText_WhenInvalid_DoesNotReturnMHeading(string text)
        {
            MNode node = new MarkdownVisitor().Visit(text);

            Assert.IsNotType<MHeading>(node);
        }

        [Theory]
        [InlineData(" Paragraph", "Paragraph", "<p>Paragraph</p>")]
        public void GivenText_WhenValid_ReturnsMParagraph(string markdown, string normalized, string html)
        {
            MNode node = new MarkdownVisitor().Visit(markdown);

            MParagraph paragraph = Assert.IsType<MParagraph>(node);
            Assert.Equal(normalized, paragraph.Text.ToString());
            Assert.Equal(normalized, node.ToString());
            Assert.Equal(html, paragraph.ToHtml().ToString());
        }

        [Theory]
        [InlineData("- item", "- item", "item", "<li>item</li>")]
        [InlineData("* item", "* item", "item", "<li>item</li>")]
        [InlineData("+ item", "+ item", "item", "<li>item</li>")]
        [InlineData("-  item", "- item", "item", "<li>item</li>")]
        [InlineData("-   item", "- item", "item", "<li>item</li>")]
        [InlineData("-    item", "- item", "item", "<li>item</li>")]
        [InlineData("  - item", "  - item", "item", "<li>item</li>")]
        [InlineData("  -  item", "  - item", "item", "<li>item</li>")]
        [InlineData("  -   item", "  - item", "item", "<li>item</li>")]
        [InlineData("  -    item", "  - item", "item", "<li>item</li>")]
        public void GivenText_WhenValid_ReturnsMBulletListItem(string markdown, string normalized, string text, string html)
        {
            MNode node = new MarkdownVisitor().Visit(markdown);

            MBulletListItem item = Assert.IsType<MBulletListItem>(node);
            Assert.Equal(normalized, node.ToString());
            Assert.Equal(text, item.Item.ToString());
            Assert.Equal(html, item.ToHtml().ToString());
        }

        [Theory]
        [InlineData("-")]
        [InlineData("- ")]
        [InlineData("-     item")]
        [InlineData("-item")]
        [InlineData("-- item")]
        public void GivenText_WhenInvalid_DoesNotReturnMBulletListItem(string text)
        {
            MNode node = new MarkdownVisitor().Visit(text);

            Assert.IsNotType<MBulletListItem>(node);
        }
    }
}