using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using AD.Xml;
using JetBrains.Annotations;

// TODO: this class throws, find a better pattern.
// ReSharper disable AssignNullToNotNullAttribute
namespace AD.OpenXml.Visits
{
    /// <summary>
    ///
    /// </summary>
    [PublicAPI]
    public static class StyleVisit
    {
        [NotNull] static readonly XNamespace W = XNamespaces.OpenXmlWordprocessingmlMain;

        [NotNull] static readonly XElement DocDefaults;
        [NotNull] static readonly XElement Normal;
        [NotNull] static readonly XElement Appendix;
        [NotNull] static readonly XElement Bibliography;
        [NotNull] static readonly XElement Box;
        [NotNull] static readonly XElement BoxCaption;
        [NotNull] static readonly XElement BoxListBullet;
        [NotNull] static readonly XElement BoxSourceNote;
        [NotNull] static readonly XElement BoxTable;
        [NotNull] static readonly XElement BlueTableBasic;
        [NotNull] static readonly XElement CaptionFigure;
        [NotNull] static readonly XElement CaptionTable;
        [NotNull] static readonly XElement Emphasis;
        [NotNull] static readonly XElement ExecutiveSummary1StParagraph;
        [NotNull] static readonly XElement ExecutiveSummaryHighlights;
        [NotNull] static readonly XElement ExecutiveSummarySidebars;
        [NotNull] static readonly XElement FigureTableSourceNote;
        [NotNull] static readonly XElement FootnoteReference;
        [NotNull] static readonly XElement FootnoteText;
        [NotNull] static readonly XElement Heading1;
        [NotNull] static readonly XElement Heading2;
        [NotNull] static readonly XElement Heading3;
        [NotNull] static readonly XElement Heading4;
        [NotNull] static readonly XElement Heading5;
        [NotNull] static readonly XElement Heading6;
        [NotNull] static readonly XElement Heading7;
        [NotNull] static readonly XElement Heading8;
        [NotNull] static readonly XElement Hyperlink;
        [NotNull] static readonly XElement ListBullet;
        [NotNull] static readonly XElement PreHeading;
        [NotNull] static readonly XElement Strong;
        [NotNull] static readonly XElement StyleNotImplemented;
        [NotNull] static readonly XElement Subscript;
        [NotNull] static readonly XElement Superscript;
        [NotNull] static readonly XElement TableOfFigures;
        [NotNull] static readonly XElement TOC1;
        [NotNull] static readonly XElement TOC2;
        [NotNull] static readonly XElement TOC3;
        [NotNull] static readonly XElement TOC4;
        [NotNull] static readonly XElement TOCHeading;

        /// <summary>
        ///
        /// </summary>
        static StyleVisit()
        {
            Assembly assembly = typeof(StyleVisit).GetTypeInfo().Assembly;

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.DocDefaults.xml"), Encoding.UTF8))
            {
                DocDefaults = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Normal.xml"), Encoding.UTF8))
            {
                Normal = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Appendix.xml"), Encoding.UTF8))
            {
                Appendix = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Bibliography.xml"), Encoding.UTF8))
            {
                Bibliography = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Box.xml"), Encoding.UTF8))
            {
                Box = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.BoxCaption.xml"), Encoding.UTF8))
            {
                BoxCaption = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.BoxListBullet.xml"), Encoding.UTF8))
            {
                BoxListBullet = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.BoxSourceNote.xml"), Encoding.UTF8))
            {
                BoxSourceNote = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.BoxTable.xml"), Encoding.UTF8))
            {
                BoxTable = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.BlueTableBasic.xml"), Encoding.UTF8))
            {
                BlueTableBasic = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.CaptionFigure.xml"), Encoding.UTF8))
            {
                CaptionFigure = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.CaptionTable.xml"), Encoding.UTF8))
            {
                CaptionTable = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Emphasis.xml"), Encoding.UTF8))
            {
                Emphasis = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.ExecutiveSummary1StParagraph.xml"), Encoding.UTF8))
            {
                ExecutiveSummary1StParagraph = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.ExecutiveSummaryHighlights.xml"), Encoding.UTF8))
            {
                ExecutiveSummaryHighlights = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.ExecutiveSummarySidebars.xml"), Encoding.UTF8))
            {
                ExecutiveSummarySidebars = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.FigureTableSourceNote.xml"), Encoding.UTF8))
            {
                FigureTableSourceNote = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.FootnoteReference.xml"), Encoding.UTF8))
            {
                FootnoteReference = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.FootnoteText.xml"), Encoding.UTF8))
            {
                FootnoteText = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Heading1.xml"), Encoding.UTF8))
            {
                Heading1 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Heading2.xml"), Encoding.UTF8))
            {
                Heading2 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Heading3.xml"), Encoding.UTF8))
            {
                Heading3 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Heading4.xml"), Encoding.UTF8))
            {
                Heading4 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Heading5.xml"), Encoding.UTF8))
            {
                Heading5 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Heading6.xml"), Encoding.UTF8))
            {
                Heading6 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Heading7.xml"), Encoding.UTF8))
            {
                Heading7 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Heading8.xml"), Encoding.UTF8))
            {
                Heading8 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Hyperlink.xml"), Encoding.UTF8))
            {
                Hyperlink = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.ListBullet.xml"), Encoding.UTF8))
            {
                ListBullet = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.PreHeading.xml"), Encoding.UTF8))
            {
                PreHeading = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Strong.xml"), Encoding.UTF8))
            {
                Strong = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.StyleNotImplemented.xml"), Encoding.UTF8))
            {
                StyleNotImplemented = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Subscript.xml"), Encoding.UTF8))
            {
                Subscript = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.Superscript.xml"), Encoding.UTF8))
            {
                Superscript = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.TableOfFigures.xml"), Encoding.UTF8))
            {
                TableOfFigures = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.TOC1.xml"), Encoding.UTF8))
            {
                TOC1 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.TOC2.xml"), Encoding.UTF8))
            {
                TOC2 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.TOC3.xml"), Encoding.UTF8))
            {
                TOC3 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.TOC4.xml"), Encoding.UTF8))
            {
                TOC4 = XElement.Parse(reader.ReadToEnd());
            }

            using (StreamReader reader = new StreamReader(assembly.GetManifestResourceStream("AD.OpenXml.Styles.TOCHeading.xml"), Encoding.UTF8))
            {
                TOCHeading = XElement.Parse(reader.ReadToEnd());
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="subject"></param>
        [NotNull]
        public static OpenXmlPackageVisitor VisitStyles([NotNull] this OpenXmlPackageVisitor subject)
        {
            if (subject is null)
                throw new ArgumentNullException(nameof(subject));

            XElement styles = Execute(subject.Styles.Clone());

            return subject.With(styles: styles);
        }

        [Pure]
        [NotNull]
        static XElement Execute([NotNull] XElement styles)
        {
            if (styles is null)
                throw new ArgumentNullException(nameof(styles));

            XElement results =
                new XElement(
                    styles.Name,
                    styles.Attributes(),
                    DocDefaults.Clone(),
                    Normal.Clone(),
                    Appendix.Clone(),
                    Bibliography.Clone(),
                    Box.Clone(),
                    BoxCaption.Clone(),
                    BoxListBullet.Clone(),
                    BoxSourceNote.Clone(),
                    BoxTable.Clone(),
                    BlueTableBasic.Clone(),
                    CaptionFigure.Clone(),
                    CaptionTable.Clone(),
                    Emphasis.Clone(),
                    ExecutiveSummary1StParagraph.Clone(),
                    ExecutiveSummaryHighlights.Clone(),
                    ExecutiveSummarySidebars.Clone(),
                    FigureTableSourceNote.Clone(),
                    FootnoteReference.Clone(),
                    FootnoteText.Clone(),
                    Heading1.Clone(),
                    Heading2.Clone(),
                    Heading3.Clone(),
                    Heading4.Clone(),
                    Heading5.Clone(),
                    Heading6.Clone(),
                    Heading7.Clone(),
                    Heading8.Clone(),
                    Hyperlink.Clone(),
                    ListBullet.Clone(),
                    PreHeading.Clone(),
                    Strong.Clone(),
                    StyleNotImplemented.Clone(),
                    Subscript.Clone(),
                    Superscript.Clone(),
                    TableOfFigures.Clone(),
                    TOC1.Clone(),
                    TOC2.Clone(),
                    TOC3.Clone(),
                    TOC4.Clone(),
                    TOCHeading.Clone());

            foreach (XElement style in results.Elements())
            {
                if (style.Elements().First().Name == W + "name")
                    continue;

                XElement name = style.Element(W + "name");
                name?.Remove();
                style.AddFirst(name);
            }

            foreach (XElement style in results.Elements())
            {
                if (style.Element(W + "name")?.Next()?.Name == W + "next")
                    continue;

                XElement next = style.Element(W + "next");
                next?.Remove();
                style.Element(W + "name")?.AddAfterSelf(next);
            }

            return results;
        }
    }
}