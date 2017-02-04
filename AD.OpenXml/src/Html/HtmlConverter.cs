//using System;
//using System.Collections.Generic;
//using System.Linq;
////using System.Web.UI;
//using System.IO;
//using System.Reflection;
//using System.Text;
//using System.Text.RegularExpressions;
//using System.Windows;
//using System.Windows.Controls;
//using System.Xml;
//using System.Xml.Xsl;
////using DocumentFormat.OpenXml;
////using DocumentFormat.OpenXml.Packaging;
////using DocumentFormat.OpenXml.Wordprocessing;
////using M = DocumentFormat.OpenXml.Math;

//namespace AD.OpenXml.Html
//{
//    internal class HtmlConverter
//    {
//        internal HtmlConverter(WordprocessingDocument document)
//        {
//            HtmlBuilder html = new HtmlBuilder(document);
//            Window window = new Window();
//            WebBrowser browser = new WebBrowser();
//            browser.NavigateToString(html.HtmlText);
//            window.Content = browser;
//            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{DateTime.Now.Ticks}.html"), html.HtmlText);
//            window.ShowDialog();
//            Environment.Exit(0);
//        }

//        private class HtmlBuilder
//        {
//            private readonly StringWriter _host;

//            private readonly HtmlTextWriter _writer;

//            internal string HtmlText { get; private set; }

//            internal HtmlBuilder(WordprocessingDocument document)
//            {
//                _host = new StringWriter();
//                _writer = new HtmlTextWriter(_host, "    ") { Indent = 1 };
//                _writer.Write(HtmlTags.DocumentType);
//                _writer.RenderBeginTag(HtmlTags.Html);
//                _writer.RenderBeginTag(HtmlTags.Head);
//                _writer.AddAttribute(HtmlAttributes.Type, "text/javascript");
//                _writer.AddAttribute(HtmlAttributes.AsyncSource, "https://cdn.mathjax.org/mathjax/latest/MathJax.js?config=MML_CHTML");
//                _writer.RenderBeginTag(HtmlTags.Script);
//                _writer.RenderEndTag();

//                PreprocessCharts(document);

//                _writer.AddAttribute(HtmlAttributes.HyperlinkReference, @"C:\Users\adren\OneDrive\Documents\Projects\USITC332.css");
//                _writer.AddAttribute(HtmlAttributes.Type, "text/css");
//                _writer.AddAttribute(HtmlAttributes.Relation, "stylesheet");
//                _writer.RenderBeginTag(HtmlTags.Link);
//                _writer.RenderEndTag();
//                _writer.AddAttribute(HtmlAttributes.CharacterSet, "UTF-8");
//                _writer.RenderBeginTag(HtmlTags.Meta);
//                _writer.RenderEndTag();
//                _writer.RenderEndTag();



//                _writer.RenderBeginTag(HtmlTags.Body);
//                Content(document);
//                _writer.RenderEndTag();

//                _writer.RenderEndTag();

//                CleansOutput();
//                _writer.Dispose();
//                _host.Dispose();
//            }

//            private void CleansOutput()
//            {
//                string a = _host.ToString();
//                a = Regex.Replace(a, @"<strong></strong>", "");
//                a = Regex.Replace(a, @"</strong><strong>", "");
//                a = Regex.Replace(a, @"<em></em>", "");
//                a = Regex.Replace(a, @"</em><em>", "");
//                a = Regex.Replace(a, @"<p></p>", "");
//                HtmlText = a;
//            }

//            private void PreprocessCharts(WordprocessingDocument document)
//            {
//                foreach (Table e in document.MainDocumentPart.Document.Body.Elements<Table>())
//                {
//                    if (e.PreviousSibling()?.GetType() != typeof(Paragraph))
//                    {
//                        continue;
//                    }
//                    if (e.PreviousSibling<Paragraph>()?.InnerText?.ToLower().Contains("columnchart") ?? false)
//                    {
//                        _writer.RenderBeginTag(HtmlTags.Script);



//                        _writer.RenderEndTag();
//                    }
//                }
//            }

//            private void Content(WordprocessingDocument document)
//            {
//                foreach (OpenXmlElement e in document.MainDocumentPart.Document.Body.Elements())
//                {
//                    switch (e.GetType().Name)
//                    {
//                        case WordElements.Paragraph:
//                        {
//                            ParagraphElements((Paragraph)e);
//                            break;
//                        }
//                        case WordElements.Table:
//                        {
//                            TableElements((Table)e);
//                            break;
//                        }
//                        case WordElements.Drawing:
//                        {
//                            DrawingElements((Drawing)e);
//                            break;
//                        }
//                        default:
//                        {
//                            break;
//                        }
//                    }
//                }
//            }

//            private void ParagraphElements(Paragraph paragraph)
//            {
//                if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.ToLower() == WordStyles.Title)
//                {
//                    FrontMatter(paragraph);
//                    return;
//                }
//                if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.ToLower() == WordStyles.SubTitle)
//                {
//                    return;
//                }
//                if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.ToLower().Contains(WordKeywords.Heading) ?? false)
//                {
//                    StyleHeading(paragraph);
//                    return;
//                }
//                if (paragraph.NextSibling()?.GetType() == typeof(Table))
//                {
//                    return;
//                }
//                if (paragraph.NextSibling()?.Descendants<Drawing>().Any() ?? false)
//                {
//                    return;
//                }
//                if (paragraph.Descendants<Drawing>().Any())
//                {
//                    DrawingElements(paragraph.Descendants<Drawing>().First());
//                    return;
//                }
//                if (paragraph.Descendants<M.OfficeMath>().Any())
//                {
//                    //foreach (M.OfficeMath math in paragraph.Descendants<M.OfficeMath>())
//                    //{
//                    //    ConvertOoxmlToMathMl(math);
//                    //}
//                    //return;
//                }
//                StyleStandardParagraph(paragraph);
//            }

//            private void FrontMatter(OpenXmlElement paragraph)
//            {
//                _writer.RenderBeginTag(HtmlTags.Header);
//                Title(paragraph);
//                SubTitle(paragraph);
//                SubTitle(paragraph);
//                _writer.RenderEndTag();
//            }

//            private void Title(OpenXmlElement paragraph)
//            {
//                _writer.AddAttribute(HtmlAttributes.Class, HtmlClasses.Title);
//                _writer.RenderBeginTag(HtmlTags.Heading1);
//                foreach (Run run in paragraph.Elements<Run>())
//                {
//                    ProcessRun(run);
//                }
//                _writer.RenderEndTag();
//            }

//            private void SubTitle(OpenXmlElement paragraph)
//            {
//                if (paragraph.NextSibling()?.GetType() == typeof(Paragraph))
//                {
//                    if (paragraph.NextSibling<Paragraph>()?.ParagraphProperties?.ParagraphStyleId?.Val?.Value.ToLower() == WordStyles.SubTitle)
//                    {
//                        Paragraph next = paragraph.NextSibling<Paragraph>();
//                        _writer.AddAttribute(HtmlAttributes.Class, HtmlClasses.SubTitle);
//                        _writer.RenderBeginTag(HtmlTags.Heading2);
//                        foreach (Run run in next.Elements<Run>())
//                        {
//                            ProcessRun(run);
//                        }
//                        _writer.RenderEndTag();
//                        next.Remove();
//                    }
//                }
//            }

//            private void TableElements(OpenXmlElement table)
//            {
//                // Start table
//                _writer.RenderBeginTag(HtmlTags.Table);
//                // Caption
//                if (table.PreviousSibling().GetType() == typeof(Paragraph))
//                {
//                    _writer.AddAttribute(HtmlAttributes.Class, _whichHeading);
//                    _writer.RenderBeginTag(HtmlTags.TableCaption);
//                    StyleCaption(table.PreviousSibling());
//                    _writer.RenderEndTag();
//                }
//                // Footer
//                if (table.NextSibling().GetType() == typeof(Paragraph))
//                {
//                    _writer.RenderBeginTag(HtmlTags.TableFooter);
//                    Paragraph p = table.NextSibling<Paragraph>();
//                    while (true)
//                    {
//                        if (p.Descendants<RunStyle>().Any(x => x.Val.Value.ToLower() == WordStyles.IntenseEmphasis) || (p.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.ToLower() == WordDeprecatedStyles.FiguresTablesSourceNote))
//                        {
//                            _writer.RenderBeginTag(HtmlTags.TableRow);
//                            _writer.AddAttribute(HtmlAttributes.ColumnSpan, table.Elements<TableRow>().First().Elements<TableCell>().Count().ToString());
//                            _writer.RenderBeginTag(HtmlTags.TableCell);
//                            foreach (Run run in p.Elements<Run>())
//                            {
//                                ProcessRun(run);
//                            }
//                            _writer.RenderEndTag();
//                            _writer.RenderEndTag();
//                            if (p.NextSibling().GetType() == typeof(Paragraph))
//                            {
//                                Paragraph nextParagraph = p.NextSibling<Paragraph>();
//                                p.Remove();
//                                p = nextParagraph;
//                            }
//                        }
//                        else
//                        {
//                            break;
//                        }
//                    }
//                    _writer.RenderEndTag();
//                }
//                // Header
//                List<bool> numericColumns = FindNumericColumns(table);
//                List<TableCell> headers = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
//                _writer.RenderBeginTag(HtmlTags.TableHeader);
//                _writer.RenderBeginTag(HtmlTags.TableRow);
//                foreach (TableCell cell in headers)
//                {
//                    _writer.AddAttribute(HtmlAttributes.Class, numericColumns[headers.IndexOf(cell)] ? HtmlClasses.Right : HtmlClasses.Left);
//                    _writer.RenderBeginTag(HtmlTags.TableHeaderCell);
//                    _writer.WriteLine(cell.Descendants<Text>()?.FirstOrDefault()?.Text?.Trim());
//                    _writer.RenderEndTag();
//                }
//                _writer.RenderEndTag();
//                _writer.RenderEndTag();
//                // Body
//                _writer.RenderBeginTag(HtmlTags.TableBody);
//                List<TableRow> rows = table.Elements<TableRow>().ToList();
//                foreach (TableRow row in rows.Skip(1))
//                {
//                    _writer.RenderBeginTag(HtmlTags.TableRow);
//                    List<TableCell> cells = row.Elements<TableCell>().ToList();
//                    foreach (TableCell cell in cells)
//                    {
//                        string attribute = numericColumns[cells.IndexOf(cell)] ? HtmlClasses.Right : HtmlClasses.Left;
//                        if ((rows.IndexOf(row) == rows.Count - 1) && row.Descendants<Text>().Any(x => x.Text?.Trim().ToLower().StartsWith(WordKeywords.Total) ?? false))
//                        {
//                            attribute = $"{attribute} {HtmlClasses.TotalRow}";
//                        }
//                        _writer.AddAttribute(HtmlAttributes.Class, attribute);
//                        _writer.RenderBeginTag(HtmlTags.TableCell);
//                        _writer.WriteLine(cell.Descendants<Text>()?.FirstOrDefault()?.Text?.Trim());
//                        _writer.RenderEndTag();
//                    }
//                    _writer.RenderEndTag();
//                }
//                _writer.RenderEndTag();
//                // End table
//                _writer.RenderEndTag();
//            }

//            private void DrawingElements(OpenXmlElement drawing)
//            {
//                _writer.RenderBeginTag(HtmlTags.Figure);
//                if (drawing.Parent.Parent.PreviousSibling().GetType() == typeof(Paragraph))
//                {
//                    _writer.AddAttribute(HtmlAttributes.Class, _whichHeading);
//                    _writer.RenderBeginTag(HtmlTags.FigureCaption);
//                    StyleCaption(drawing.Parent.Parent.PreviousSibling());
//                    _writer.RenderEndTag();
//                }
//                //List<bool> numericColumns = FindNumericColumns(table);
//                //List<TableCell> headers = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
//                //foreach (TableCell cell in headers)
//                //{
//                //_writer.AddAttribute(HtmlAttributes.Class, numericColumns[headers.IndexOf(cell)] ? "right" : "left");
//                //_writer.RenderBeginTag(HtmlTags.TableHeader);
//                //_writer.WriteLine(cell.Descendants<Text>()?.FirstOrDefault()?.Text?.Trim());
//                //_writer.RenderEndTag();
//                //}
//                //foreach (TableRow row in table.Elements<TableRow>().Skip(1))
//                //{
//                //    _writer.RenderBeginTag(HtmlTags.TableRow);
//                //    List<TableCell> cells = row.Elements<TableCell>().ToList();
//                //    foreach (TableCell cell in cells)
//                //    {
//                //        _writer.AddAttribute(HtmlAttributes.Class, numericColumns[cells.IndexOf(cell)] ? "right" : "left");
//                //        _writer.RenderBeginTag(HtmlTags.TableCell);
//                //        _writer.WriteLine(cell.Descendants<Text>()?.FirstOrDefault()?.Text?.Trim());
//                //        _writer.RenderEndTag();
//                //    }
//                //    _writer.RenderEndTag();
//                //}
//                _writer.RenderEndTag();
//            }

//            private static List<bool> FindNumericColumns(OpenXmlElement table)
//            {
//                List<bool> numeric = new List<bool>();
//                List<TableRow> rows = table.Elements<TableRow>().Skip(1).ToList();
//                int cells = rows[0].Elements<TableCell>().Count();
//                for (int i = 0; i < cells; i++)
//                {
//                    List<bool> test = rows.Select(row => Regex.IsMatch(row?.Elements<TableCell>()?.ElementAt(i)?.Descendants<Text>()?.FirstOrDefault()?.Text?.Trim() ?? string.Empty, RegularExpressions.NumbersCommasDecimals)).ToList();
//                    numeric.Add(test.All(x => x));
//                }
//                return numeric;
//            }

//            private void StyleHeading(Paragraph paragraph)
//            {
//                OpenHeadingLevel(paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
//                foreach (Run r in paragraph.Elements<Run>())
//                {
//                    ProcessRun(r);
//                }
//                CloseHeadingLevel(paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
//            }

//            private string _whichHeading;

//            private void OpenHeadingLevel(string value)
//            {
//                switch (value.ToLower().Substring(0, 8))
//                {
//                    case WordStyles.Heading1:
//                    {
//                        _whichHeading = HtmlClasses.Chapter;
//                        _writer.AddAttribute(HtmlAttributes.Class, _whichHeading);
//                        _writer.RenderBeginTag(HtmlTags.Heading1);
//                        break;
//                    }
//                    case WordStyles.Heading2:
//                    {
//                        _writer.RenderBeginTag(HtmlTags.Heading2);
//                        break;
//                    }
//                    case WordStyles.Heading3:
//                    {
//                        _writer.RenderBeginTag(HtmlTags.Heading3);
//                        break;
//                    }
//                    case WordStyles.Heading4:
//                    {
//                        _writer.RenderBeginTag(HtmlTags.Heading4);
//                        break;
//                    }
//                    case WordStyles.Heading5:
//                    {
//                        _writer.RenderBeginTag(HtmlTags.Heading5);
//                        break;
//                    }
//                    case WordStyles.Heading6:
//                    {
//                        _writer.RenderBeginTag(HtmlTags.Heading6);
//                        break;
//                    }
//                    case WordStyles.Heading7:
//                    {
//                        _whichHeading = HtmlClasses.Appendix;
//                        _writer.AddAttribute(HtmlAttributes.Class, _whichHeading);
//                        _writer.RenderBeginTag(HtmlTags.Heading1);
//                        break;
//                    }
//                    default:
//                    {
//                        break;
//                    }
//                }
//            }

//            private void CloseHeadingLevel(string value)
//            {
//                switch (value.ToLower().Substring(0, 8))
//                {
//                    case WordStyles.Heading1:
//                    {
//                        _writer.RenderEndTag();
//                        break;
//                    }
//                    case WordStyles.Heading2:
//                    {
//                        _writer.RenderEndTag();
//                        break;
//                    }
//                    case WordStyles.Heading3:
//                    {
//                        _writer.RenderEndTag();
//                        break;
//                    }
//                    case WordStyles.Heading4:
//                    {
//                        _writer.RenderEndTag();
//                        break;
//                    }
//                    case WordStyles.Heading5:
//                    {
//                        _writer.RenderEndTag();
//                        break;
//                    }
//                    case WordStyles.Heading6:
//                    {
//                        _writer.RenderEndTag();
//                        break;
//                    }
//                    case WordStyles.Heading7:
//                    {
//                        _writer.RenderEndTag();
//                        break;
//                    }
//                    default:
//                    {
//                        break;
//                    }
//                }
//            }

//            private void StyleCaption(OpenXmlElement paragraph)
//            {
//                string text = null;
//                foreach (Run r in paragraph.Elements<Run>())
//                {
//                    if (r.Descendants<FieldChar>().Any())
//                    {
//                        continue;
//                    }
//                    if (r.Elements<Text>().FirstOrDefault()?.Text?.ToLower().Contains(WordKeywords.Sequence) ?? false)
//                    {
//                        continue;
//                    }
//                    if (r.Elements<Text>().FirstOrDefault()?.Text?.ToLower().Contains(WordKeywords.StyleReference) ?? false)
//                    {
//                        continue;
//                    }
//                    foreach (Text t in r.Elements<Text>())
//                    {
//                        text += t.Text;
//                    }
//                }
//                Regex regex = new Regex(RegularExpressions.CaptionIndex);
//                _writer.Write(regex.Replace(text ?? string.Empty, string.Empty));
//            }

//            private void StyleStandardParagraph(OpenXmlElement paragraph)
//            {
//                if (!paragraph.Descendants<Text>().Any())
//                {
//                    return;
//                }
//                _writer.RenderBeginTag(HtmlTags.Paragraph);
//                foreach (OpenXmlElement e in paragraph.Elements())
//                {
//                    if (e.GetType() == typeof(Run))
//                    {
//                        ProcessRun((Run)e);
//                    }
//                    else if (e.GetType() == typeof(M.OfficeMath))
//                    {
//                        ConvertOoxmlToMathMl(e);
//                    }
//                }
//                //foreach (Run r in paragraph.Elements<Run>())
//                //{
//                //    ProcessRun(r);
//                //}
//                _writer.RenderEndTag();
//            }

//            private void ProcessRun(Run run)
//            {
//                if ((run.RunProperties?.Bold != null) || (run.RunProperties?.RunStyle?.Val?.Value?.ToLower() == WordStyles.Strong))
//                {
//                    StyleBold(run);
//                    return;
//                }
//                if ((run.RunProperties?.Italic != null) || (run.RunProperties?.RunStyle?.Val?.Value?.ToLower() == WordStyles.Emphasis))
//                {
//                    StyleItalics(run);
//                    return;
//                }
//                if ((run.RunProperties?.RunStyle?.Val?.Value?.ToLower() == WordStyles.SuperScript) || (run.RunProperties?.RunStyle?.Elements<VerticalTextAlignment>().FirstOrDefault()?.Val?.Value == VerticalPositionValues.Superscript))
//                {
//                    StyleSuperscript(run);
//                    return;
//                }
//                if ((run.RunProperties?.RunStyle?.Val?.Value?.ToLower() == WordStyles.Subscript) || (run.RunProperties?.RunStyle?.Elements<VerticalTextAlignment>().FirstOrDefault()?.Val?.Value == VerticalPositionValues.Subscript))
//                {
//                    StyleSubscript(run);
//                    return;
//                }
//                foreach (Text t in run.Elements<Text>())
//                {
//                    _writer.Write(t.Text);
//                }
//            }

//            private void StyleBold(OpenXmlElement run)
//            {
//                if (!run.Elements<Text>().Any())
//                {
//                    return;
//                }
//                _writer.RenderBeginTag(HtmlTags.Strong);
//                foreach (Text t in run.Elements<Text>())
//                {
//                    _writer.Write(t.Text);
//                }
//                _writer.RenderEndTag();
//            }

//            private void StyleItalics(OpenXmlElement run)
//            {
//                if (!run.Elements<Text>().Any())
//                {
//                    return;
//                }
//                _writer.RenderBeginTag(HtmlTags.Emphasis);
//                foreach (Text t in run.Elements<Text>())
//                {
//                    _writer.Write(t.Text);
//                }
//                _writer.RenderEndTag();
//            }

//            private void StyleSuperscript(OpenXmlElement run)
//            {
//                if (!run.Elements<Text>().Any())
//                {
//                    return;
//                }
//                _writer.RenderBeginTag(HtmlTags.Superscript);
//                foreach (Text t in run.Elements<Text>())
//                {
//                    _writer.Write(t.Text);
//                }
//                _writer.RenderEndTag();
//            }

//            private void StyleSubscript(OpenXmlElement run)
//            {
//                if (!run.Elements<Text>().Any())
//                {
//                    return;
//                }
//                _writer.RenderBeginTag(HtmlTags.Subscript);
//                foreach (Text t in run.Elements<Text>())
//                {
//                    _writer.Write(t.Text);
//                }
//                _writer.RenderEndTag();
//            }

//            private void ConvertOoxmlToMathMl(OpenXmlElement math)
//            {
//                Assembly assembly = Assembly.GetExecutingAssembly();
//                using (Stream stream = assembly.GetManifestResourceStream("Getdata_document_assembler.Resources.OMML2MML.XSL"))
//                {
//                    if (stream == null)
//                    {
//                        return;
//                    }
//                    using (XmlReader reader = XmlReader.Create(stream))
//                    {
//                        string tempFile = Path.GetTempFileName();
//                        string tempFile2 = Path.GetTempFileName();
//                        using (XmlTextWriter writer = new XmlTextWriter(tempFile, new UTF8Encoding()))
//                        {
//                            math.WriteTo(writer);
//                        }
//                        XslCompiledTransform xslt = new XslCompiledTransform();
//                        xslt.Load(reader);
//                        xslt.Transform(tempFile, tempFile2);
//                        string text = File.ReadAllText(tempFile2);
//                        text = text.Replace("mml:", null);
//                        text = text.Replace(@"xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math""", null);
//                        text = text.Replace(@"<?xml version=""1.0"" encoding=""utf-16""?>", null);
//                        _writer.Write(text);
//                    }
//                }
//            }
//        }

//        private static class HtmlTags
//        {
//            public const string Body = "body";
//            public const string Break = "br";
//            public const string DocumentType = "<!DOCTYPE html>";
//            public const string Emphasis = "em";
//            public const string Figure = "figure";
//            public const string FigureCaption = "figcaption";
//            public const string Heading1 = "h1";
//            public const string Heading2 = "h2";
//            public const string Heading3 = "h3";
//            public const string Heading4 = "h4";
//            public const string Heading5 = "h5";
//            public const string Heading6 = "h6";
//            public const string Head = "head";
//            public const string Header = "header";
//            public const string Html = "html";
//            public const string Image = "img";
//            public const string Link = "link";
//            public const string Meta = "meta";
//            public const string Paragraph = "p";
//            public const string Script = "script";
//            public const string Strong = "strong";
//            public const string Subscript = "sub";
//            public const string Superscript = "sup";
//            public const string TableCaption = "caption";
//            public const string Table = "table";
//            public const string TableBody = "tbody";
//            public const string TableFooter = "tfoot";
//            public const string TableHeader = "thead";
//            public const string TableHeaderCell = "th";
//            public const string TableRow = "tr";
//            public const string TableCell = "td";
//        }

//        private static class HtmlAttributes
//        {
//            public const string AltText = "alt";
//            public const string Class = "class";
//            public const string ColumnSpan = "colspan";
//            public const string Type = "type";
//            public const string AsyncSource = "async src";
//            public const string Source = "src";
//            public const string CharacterSet = "charset";
//            public const string HyperlinkReference = "href";
//            public const string Relation = "rel";
//        }

//        private static class HtmlClasses
//        {
//            public const string Chapter = "chapter";
//            public const string Appendix = "appendix";
//            public const string Left = "left";
//            public const string Right = "right";
//            public const string TotalRow = "totalrow";
//            public const string Title = "title";
//            public const string SubTitle = "subtitle";
//        }

//        private static class WordStyles
//        {
//            public const string Heading1 = "heading1";
//            public const string Heading2 = "heading2";
//            public const string Heading3 = "heading3";
//            public const string Heading4 = "heading4";
//            public const string Heading5 = "heading5";
//            public const string Heading6 = "heading6";
//            public const string Heading7 = "heading7";
//            public const string SubtleEmphasis = "subtleemphasis";
//            public const string Title = "title";
//            public const string SubTitle = "subtitle";
//            public const string IntenseEmphasis = "intenseemphasis";
//            public const string Subscript = "subscript";
//            public const string SuperScript = "superscript";
//            public const string Emphasis = "emphasis";
//            public const string Strong = "strong";
//        }

//        private static class WordElements
//        {
//            public const string Drawing = "Drawing";
//            public const string Math = "OfficeMath";
//            public const string Paragraph = "Paragraph";
//            public const string Table = "Table";
//        }

//        private static class WordKeywords
//        {
//            public const string Heading = "heading";
//            public const string Total = "total";
//            public const string Sequence = "seq";
//            public const string StyleReference = "styleref";
//        }

//        private static class WordDeprecatedStyles
//        {
//            public const string FiguresTablesSourceNote = "figurestablessourcenote";
//        }

//        private static class RegularExpressions
//        {
//            public const string NumbersCommasDecimals = @"(^-*[0-9]{1,3}(,[0-9]{3})*(\.[0-9]+)?$|^$)";
//            public const string CaptionIndex = @"((Table|Figure) [0-9A-Z]+\.[0-9]+)";
//        }
//    }
//}
