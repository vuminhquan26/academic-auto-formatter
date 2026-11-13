using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Draw = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;

namespace AutoFormatter
{
    public class DocxEngine
    {
        public static void FormatDocument(string filePath, FormatConfig config)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null) return;
                var body = mainPart.Document.Body;

                var settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
                if (settingsPart.Settings == null) settingsPart.Settings = new Settings();

                var updateFieldsList = settingsPart.Settings.Elements<UpdateFieldsOnOpen>().ToList();
                foreach (var item in updateFieldsList) { item.Remove(); }
                settingsPart.Settings.AppendChild(new UpdateFieldsOnOpen() { Val = false });

                uint leftTwips = (uint)(config.PageSetup.LeftCm * 567);
                uint rightTwips = (uint)(config.PageSetup.RightCm * 567);
                uint topTwips = (uint)(config.PageSetup.TopCm * 567);
                uint bottomTwips = (uint)(config.PageSetup.BottomCm * 567);

                var sectPrs = body.Descendants<SectionProperties>().ToList();
                if (!sectPrs.Any()) { sectPrs.Add(new SectionProperties()); body.AppendChild(sectPrs.First()); }

                foreach (var sectPr in sectPrs)
                {
                    var pageMargin = sectPr.Elements<PageMargin>().FirstOrDefault() ?? sectPr.AppendChild(new PageMargin());
                    pageMargin.Left = leftTwips;
                    pageMargin.Right = rightTwips;
                    pageMargin.Top = (int)topTwips;
                    pageMargin.Bottom = (int)bottomTwips;
                }

                var mirrorMargins = settingsPart.Settings.Elements<MirrorMargins>().FirstOrDefault();
                if (config.PageSetup.IsTwoSidedPrint) { if (mirrorMargins == null) settingsPart.Settings.AppendChild(new MirrorMargins()); }
                else { mirrorMargins?.Remove(); }

                string targetFont = config.Typography.FontName;
                string targetSize = (config.Typography.FontSize * 2).ToString();

                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    var pPr = paragraph.ParagraphProperties ?? paragraph.AppendChild(new ParagraphProperties());

                    var jc = pPr.Elements<Justification>().FirstOrDefault();
                    if (jc == null) { jc = new Justification(); pPr.AppendChild(jc); }

                    if (config.Typography.EnforceLineSpacing)
                    {
                        var spacing = pPr.Elements<SpacingBetweenLines>().FirstOrDefault() ?? pPr.AppendChild(new SpacingBetweenLines());
                        int lineTwips = (int)Math.Round(config.Typography.LineSpacingValue * 240);
                        spacing.Line = lineTwips.ToString();
                        spacing.LineRule = LineSpacingRuleValues.Auto;
                    }

                    string paraText = paragraph.InnerText.Trim();
                    string upperText = paraText.ToUpper();

                    if (config.Automation.FixReferenceIndentation && Regex.IsMatch(paraText, @"^\[\d+\]\s"))
                    {
                        var ind = pPr.Elements<Indentation>().FirstOrDefault() ?? pPr.AppendChild(new Indentation());
                        ind.FirstLine = "0"; ind.Hanging = "0"; ind.Left = "0";
                    }

                    bool isHeading1 = false, isHeading2 = false, isHeading3 = false;

                    if (config.Typography.AutoFormatHeadings)
                    {
                        isHeading1 = paraText.Length < 100 && (
                                     upperText.StartsWith("PHẦN ") ||
                                     upperText.StartsWith("CHƯƠNG ") ||
                                     upperText.StartsWith("MỞ ĐẦU") ||
                                     upperText.StartsWith("LỜI MỞ ĐẦU") ||
                                     upperText.StartsWith("LỜI CẢM ƠN") ||
                                     upperText.StartsWith("TÓM TẮT") ||
                                     upperText.StartsWith("KẾT LUẬN") ||
                                     upperText.StartsWith("MỤC LỤC") ||
                                     upperText.StartsWith("DANH MỤC") ||
                                     upperText.StartsWith("TÀI LIỆU THAM KHẢO") ||
                                     upperText.StartsWith("PHỤ LỤC"));

                        isHeading2 = Regex.IsMatch(paraText, @"^\d{1,2}\.\d{1,2}\.?\s+\p{L}");
                        isHeading3 = Regex.IsMatch(paraText, @"^\d{1,2}\.\d{1,2}\.\d{1,2}\.?\s+\p{L}");

                        bool isHeading = isHeading1 || isHeading2 || isHeading3;

                        if (isHeading)
                        {
                            var pStyle = pPr.Elements<ParagraphStyleId>().FirstOrDefault() ?? pPr.AppendChild(new ParagraphStyleId());
                            if (isHeading1) pStyle.Val = "Heading1";
                            else if (isHeading2) pStyle.Val = "Heading2";
                            else if (isHeading3) pStyle.Val = "Heading3";

                            if (upperText.StartsWith("PHẦN ") || upperText.StartsWith("CHƯƠNG ")) jc.Val = JustificationValues.Left;
                            else if (isHeading1) jc.Val = JustificationValues.Center;
                            else jc.Val = JustificationValues.Left;
                        }
                        else
                        {
                            if (jc.Val == null || (jc.Val != JustificationValues.Center && jc.Val != JustificationValues.Right))
                            {
                                jc.Val = JustificationValues.Both;
                            }
                        }

                        foreach (var run in paragraph.Descendants<Run>())
                        {
                            var rPr = run.RunProperties ?? run.AppendChild(new RunProperties());
                            var runFonts = rPr.Elements<RunFonts>().FirstOrDefault() ?? rPr.AppendChild(new RunFonts { Ascii = targetFont, HighAnsi = targetFont, ComplexScript = targetFont });
                            runFonts.Ascii = targetFont; runFonts.HighAnsi = targetFont; runFonts.ComplexScript = targetFont;

                            var fontSize = rPr.Elements<FontSize>().FirstOrDefault() ?? rPr.AppendChild(new FontSize());
                            var bold = rPr.Elements<Bold>().FirstOrDefault();

                            if (isHeading1) { fontSize.Val = "28"; if (bold == null) rPr.AppendChild(new Bold()); }
                            else if (isHeading2) { fontSize.Val = targetSize; if (bold == null) rPr.AppendChild(new Bold()); }
                            else if (isHeading3) { fontSize.Val = targetSize; if (bold == null) rPr.AppendChild(new Bold()); }
                            else { fontSize.Val = targetSize; }
                        }
                    }
                }

                if (config.Automation.FixTableAndFigureWidth)
                {
                    foreach (var table in body.Descendants<Table>())
                    {
                        if (table.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any()) continue;

                        bool isCoverTable = false;
                        var grid = table.Elements<TableGrid>().FirstOrDefault();
                        if (grid != null && grid.Elements<GridColumn>().Any(c => c.Width != null && (c.Width.Value == "2500" || c.Width.Value == "3800")))
                            isCoverTable = true;

                        if (isCoverTable) continue;

                        var tblPr = table.Elements<TableProperties>().FirstOrDefault();
                        if (tblPr == null) { tblPr = new TableProperties(); table.InsertAt(tblPr, 0); }

                        var tableJc = tblPr.Elements<TableJustification>().FirstOrDefault() ?? tblPr.AppendChild(new TableJustification());
                        tableJc.Val = TableRowAlignmentValues.Center;

                        var tblWidth = tblPr.Elements<TableWidth>().FirstOrDefault() ?? tblPr.AppendChild(new TableWidth());
                        tblWidth.Type = TableWidthUnitValues.Pct; tblWidth.Width = "3750";
                    }

                    var drawings = body.Descendants<Drawing>().ToList();
                    foreach (var drawing in drawings)
                    {
                        var parentParagraph = drawing.Ancestors<Paragraph>().FirstOrDefault();
                        if (parentParagraph != null)
                        {
                            var pPr = parentParagraph.ParagraphProperties ?? parentParagraph.AppendChild(new ParagraphProperties());

                            var ind = pPr.Elements<Indentation>().FirstOrDefault();
                            if (ind != null) { ind.FirstLine = "0"; ind.Left = "0"; ind.Right = "0"; }

                            var docJc = pPr.Elements<Justification>().FirstOrDefault() ?? pPr.AppendChild(new Justification());
                            docJc.Val = JustificationValues.Center;

                            var anchor = drawing.Elements<Wp.Anchor>().FirstOrDefault();
                            if (anchor != null)
                            {
                                var inline = ConvertAnchorToInline(anchor);
                                drawing.AppendChild(inline);
                                anchor.Remove();
                            }

                            bool isLogo = false;
                            var docPr = drawing.Descendants<Wp.DocProperties>().FirstOrDefault();
                            if (docPr != null && docPr.Name != null && docPr.Name.Value == "CoverLogo") isLogo = true;

                            if (!isLogo)
                            {
                                var extent = drawing.Descendants<Wp.Extent>().FirstOrDefault();
                                if (extent != null && extent.Cx != null && extent.Cy != null)
                                {
                                    long maxWidthEmus = 4185000;
                                    if (extent.Cx.Value > maxWidthEmus)
                                    {
                                        double ratio = (double)maxWidthEmus / extent.Cx.Value;
                                        extent.Cx = maxWidthEmus;
                                        extent.Cy = (long)(extent.Cy.Value * ratio);

                                        var aExtents = drawing.Descendants<Draw.Extents>().ToList();
                                        foreach (var aExt in aExtents) { aExt.Cx = extent.Cx; aExt.Cy = extent.Cy; }
                                    }
                                }
                            }
                        }
                    }
                }

                if (config.Automation.FixEquations)
                {
                    var eqParagraphs = body.Elements<Paragraph>().Where(p => p.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any()).ToList();

                    foreach (var originalEqPara in eqParagraphs)
                    {
                        Table eqTable = CreateEquationTable(originalEqPara);
                        body.InsertBefore(eqTable, originalEqPara);
                        originalEqPara.Remove();
                    }
                }

                List<OpenXmlElement> elementsToInsertAtTop = new List<OpenXmlElement>();

                if (config.Cover.GenerateCover)
                {
                    // Section đầu tiên của file gốc (sẽ trở thành section thứ 2 sau khi chèn bìa)
                    var firstOriginalSectPr = body.Descendants<SectionProperties>().FirstOrDefault();

                    CreateCoverPage(elementsToInsertAtTop, mainPart, config.Cover, targetFont);

                    if (config.Cover.AddPageBorder)
                    {
                        // Bìa là một SECTION RIÊNG: gắn sectPr (có viền) vào đoạn cuối cùng của bìa.
                        // Làm vậy không sinh thêm đoạn trống nên bìa không bị tràn sang trang trắng.
                        var lastCoverPara = elementsToInsertAtTop.OfType<Paragraph>().LastOrDefault();
                        if (lastCoverPara == null)
                        {
                            lastCoverPara = new Paragraph();
                            elementsToInsertAtTop.Add(lastCoverPara);
                        }

                        var lastPPr = lastCoverPara.ParagraphProperties ?? lastCoverPara.PrependChild(new ParagraphProperties());
                        lastPPr.AppendChild(CreateCoverSectionProperties(body, leftTwips, rightTwips, topTwips, bottomTwips));

                        // Đảm bảo section sau bìa bắt đầu ở trang mới (không bị "continuous")
                        var nextType = firstOriginalSectPr?.Elements<SectionType>().FirstOrDefault();
                        if (nextType != null) nextType.Val = SectionMarkValues.NextPage;
                    }
                    else
                    {
                        elementsToInsertAtTop.Add(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    }
                }

                if (config.Toc.GenerateTOC || config.Toc.GenerateTOF || config.Toc.GenerateTOT)
                {
                    elementsToInsertAtTop.Add(new Paragraph(new Run(new Text("👉 Vui lòng nhấn Ctrl+A rồi nhấn F9 để nạp Mục Lục")) { RunProperties = new RunProperties(new Italic(), new Color() { Val = "808080" }) }));
                }

                if (config.Toc.GenerateTOT)
                {
                    elementsToInsertAtTop.Add(CreateHeading("DANH MỤC BẢNG BIỂU", targetFont));
                    elementsToInsertAtTop.Add(CreateFieldCode(@" TOC \h \z \c ""Bảng"" "));
                    elementsToInsertAtTop.Add(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }

                if (config.Toc.GenerateTOF)
                {
                    elementsToInsertAtTop.Add(CreateHeading("DANH MỤC HÌNH VẼ", targetFont));
                    elementsToInsertAtTop.Add(CreateFieldCode(@" TOC \h \z \c ""Hình"" "));
                    elementsToInsertAtTop.Add(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }

                if (config.Toc.GenerateTOC)
                {
                    elementsToInsertAtTop.Add(CreateHeading("MỤC LỤC", targetFont));
                    elementsToInsertAtTop.Add(CreateFieldCode(@" TOC \o ""1-3"" \h \z \u "));
                    elementsToInsertAtTop.Add(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                }

                int insertIndex = 0;
                foreach (var elem in elementsToInsertAtTop)
                {
                    body.InsertAt(elem, insertIndex);
                    insertIndex++;
                }

                mainPart.Document.Save();
            }
        }

        /// <summary>
        /// Tạo SectionProperties cho section trang bìa: khổ giấy + lề + khung viền đôi.
        /// Thứ tự phần tử đúng schema: pgSz -> pgMar -> pgBorders.
        /// </summary>
        private static SectionProperties CreateCoverSectionProperties(Body body, uint left, uint right, uint top, uint bottom)
        {
            var sectPr = new SectionProperties();

            // Lấy khổ giấy từ section gốc nếu có, mặc định A4
            var srcPageSize = body.Descendants<SectionProperties>()
                                  .Select(s => s.Elements<PageSize>().FirstOrDefault())
                                  .FirstOrDefault(p => p != null);
            if (srcPageSize != null) sectPr.AppendChild((PageSize)srcPageSize.CloneNode(true));
            else sectPr.AppendChild(new PageSize() { Width = 11906U, Height = 16838U });

            sectPr.AppendChild(new PageMargin()
            {
                Left = left,
                Right = right,
                Top = (int)top,
                Bottom = (int)bottom,
                Header = 708U,
                Footer = 708U,
                Gutter = 0U
            });

            sectPr.AppendChild(new PageBorders(
                new TopBorder() { Val = BorderValues.Double, Size = 18, Space = 24, Color = "000080" },
                new LeftBorder() { Val = BorderValues.Double, Size = 18, Space = 24, Color = "000080" },
                new BottomBorder() { Val = BorderValues.Double, Size = 18, Space = 24, Color = "000080" },
                new RightBorder() { Val = BorderValues.Double, Size = 18, Space = 24, Color = "000080" }
            )
            {
                Display = PageBorderDisplayValues.AllPages,   // section này chỉ chứa trang bìa
                OffsetFrom = PageBorderOffsetValues.Text      // Measure from: Text (khung ôm theo vùng chữ)
            });

            return sectPr;
        }

        private static void CreateCoverPage(List<OpenXmlElement> elements, MainDocumentPart mainPart, CoverGeneratorConfig coverConfig, string fontName)
        {
            if (!string.IsNullOrEmpty(coverConfig.UniversityName))
                elements.Add(CreateCoverText(coverConfig.UniversityName.ToUpper(), fontName, "28", true));
            if (!string.IsNullOrEmpty(coverConfig.FacultyName))
                elements.Add(CreateCoverText(coverConfig.FacultyName.ToUpper(), fontName, "28", true));

            elements.Add(new Paragraph(new Run(new Text(""))));
            elements.Add(new Paragraph(new Run(new Text(""))));

            if (!string.IsNullOrEmpty(coverConfig.LogoPath) && File.Exists(coverConfig.LogoPath))
            {
                try
                {
                    byte[] imageArray = File.ReadAllBytes(coverConfig.LogoPath);
                    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                    using (MemoryStream ms = new MemoryStream(imageArray))
                    {
                        imagePart.FeedData(ms);
                    }
                    string relationshipId = mainPart.GetIdOfPart(imagePart);
                    var imgElement = CreateImageElement(relationshipId, 150, 150);
                    var pImg = new Paragraph(new Run(imgElement)) { ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center }) };
                    elements.Add(pImg);
                }
                catch { }
            }

            elements.Add(new Paragraph(new Run(new Text(""))));
            elements.Add(new Paragraph(new Run(new Text(""))));

            if (!string.IsNullOrEmpty(coverConfig.ReportType))
                elements.Add(CreateCoverText(coverConfig.ReportType.ToUpper(), fontName, "32", true));

            elements.Add(new Paragraph(new Run(new Text(""))));

            if (!string.IsNullOrEmpty(coverConfig.TopicName))
                elements.Add(CreateCoverText(coverConfig.TopicName, fontName, "36", true));

            for (int i = 0; i < 4; i++) elements.Add(new Paragraph(new Run(new Text(""))));

            Table infoTable = new Table();
            TableProperties tblPr = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder() { Val = BorderValues.None, Size = 0 },
                    new BottomBorder() { Val = BorderValues.None, Size = 0 },
                    new LeftBorder() { Val = BorderValues.None, Size = 0 },
                    new RightBorder() { Val = BorderValues.None, Size = 0 },
                    new InsideHorizontalBorder() { Val = BorderValues.None, Size = 0 },
                    new InsideVerticalBorder() { Val = BorderValues.None, Size = 0 }
                ),
                new TableJustification() { Val = TableRowAlignmentValues.Center }
            );
            infoTable.AppendChild(tblPr);

            TableGrid tblGrid = new TableGrid(
                new GridColumn() { Width = "3800" },
                new GridColumn() { Width = "3500" },
                new GridColumn() { Width = "1500" }
            );
            infoTable.AppendChild(tblGrid);

            if (!string.IsNullOrEmpty(coverConfig.InstructorName))
                infoTable.AppendChild(CreateCoverTableRow3Cols("Giảng viên hướng dẫn:", coverConfig.InstructorName, "", fontName, true));

            foreach (var field in coverConfig.CommonFields)
            {
                infoTable.AppendChild(CreateCoverTableRow3Cols(field.Label, field.Value, "", fontName, true));
            }

            if (coverConfig.Students != null && coverConfig.Students.Count > 0)
            {
                if (coverConfig.IsGroupProject)
                {
                    string groupText = !string.IsNullOrEmpty(coverConfig.GroupName) ? $"Nhóm {coverConfig.GroupName}" : "";
                    infoTable.AppendChild(CreateCoverTableRow3Cols("Nhóm sinh viên thực hiện:", groupText, "", fontName, true));

                    for (int i = 0; i < coverConfig.Students.Count; i++)
                    {
                        var stu = coverConfig.Students[i];
                        infoTable.AppendChild(CreateCoverTableRow3Cols("", $"{i + 1}. {stu.Name}", stu.StudentId, fontName, true));
                    }
                }
                else
                {
                    var stu = coverConfig.Students[0];
                    if (!string.IsNullOrEmpty(stu.Name))
                        infoTable.AppendChild(CreateCoverTableRow3Cols("Sinh viên thực hiện:", stu.Name, stu.StudentId, fontName, true));
                }
            }
            elements.Add(infoTable);

            for (int i = 0; i < 7; i++) elements.Add(new Paragraph(new Run(new Text(""))));

            if (!string.IsNullOrEmpty(coverConfig.LocationAndDate))
            {
                var pLoc = new Paragraph(new Run(new Text(coverConfig.LocationAndDate)) { RunProperties = new RunProperties(new RunFonts { Ascii = fontName, HighAnsi = fontName }, new Italic(), new FontSize() { Val = "26" }) })
                {
                    ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
                };
                elements.Add(pLoc);
            }
        }

        private static Paragraph CreateCoverText(string text, string fontName, string fontSize, bool isBold)
        {
            var runPr = new RunProperties(new RunFonts { Ascii = fontName, HighAnsi = fontName }, new FontSize() { Val = fontSize });
            if (isBold) runPr.AppendChild(new Bold());

            return new Paragraph(new Run(new Text(text)) { RunProperties = runPr })
            {
                ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
            };
        }

        private static TableRow CreateCoverTableRow3Cols(string col1Text, string col2Text, string col3Text, string fontName, bool isBoldValue)
        {
            TableRow tr = new TableRow();

            var noBorderCellPr = new TableCellBorders(
                new TopBorder() { Val = BorderValues.None },
                new BottomBorder() { Val = BorderValues.None },
                new LeftBorder() { Val = BorderValues.None },
                new RightBorder() { Val = BorderValues.None }
            );

            var pLabel = new Paragraph();
            if (!string.IsNullOrEmpty(col1Text))
                pLabel.AppendChild(new Run(new Text(col1Text)) { RunProperties = new RunProperties(new RunFonts { Ascii = fontName, HighAnsi = fontName }, new FontSize() { Val = "26" }, new Bold()) });
            pLabel.ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Right });
            TableCell tc1 = new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3800" }), pLabel);
            tc1.TableCellProperties.AppendChild(noBorderCellPr.CloneNode(true));
            tr.Append(tc1);

            var runPr2 = new RunProperties(new RunFonts { Ascii = fontName, HighAnsi = fontName }, new FontSize() { Val = "26" });
            if (isBoldValue) runPr2.AppendChild(new Bold());
            var pName = new Paragraph(new Run(new Text(col2Text)) { RunProperties = runPr2 })
            {
                ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Left })
            };
            pName.ParagraphProperties.AppendChild(new Indentation() { Left = "200" });

            if (string.IsNullOrEmpty(col3Text))
            {
                TableCell tc2 = new TableCell(
                    new TableCellProperties(
                        new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "5000" },
                        new GridSpan() { Val = 2 }
                    ), pName);
                tc2.TableCellProperties.AppendChild(noBorderCellPr.CloneNode(true));
                tr.Append(tc2);
            }
            else
            {
                TableCell tc2 = new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "3500" }), pName);
                tc2.TableCellProperties.AppendChild(noBorderCellPr.CloneNode(true));
                tr.Append(tc2);

                var runPr3 = new RunProperties(new RunFonts { Ascii = fontName, HighAnsi = fontName }, new FontSize() { Val = "26" });
                if (isBoldValue) runPr3.AppendChild(new Bold());
                var pId = new Paragraph(new Run(new Text(col3Text)) { RunProperties = runPr3 })
                {
                    ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Left })
                };
                TableCell tc3 = new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "1500" }), pId);
                tc3.TableCellProperties.AppendChild(noBorderCellPr.CloneNode(true));
                tr.Append(tc3);
            }

            return tr;
        }

        private static Drawing CreateImageElement(string relationshipId, long widthPx, long heightPx)
        {
            long emuWidth = widthPx * 9525;
            long emuHeight = heightPx * 9525;

            var element =
                new Drawing(
                    new Wp.Inline(
                        new Wp.Extent() { Cx = emuWidth, Cy = emuHeight },
                        new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                        new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "CoverLogo" },
                        new Wp.NonVisualGraphicFrameDrawingProperties(new Draw.GraphicFrameLocks() { NoChangeAspect = true }),
                        new Draw.Graphic(
                            new Draw.GraphicData(
                                new Pic.Picture(
                                    new Pic.NonVisualPictureProperties(
                                        new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Logo.png" },
                                        new Pic.NonVisualPictureDrawingProperties()),
                                    new Pic.BlipFill(
                                        new Draw.Blip(new Draw.BlipExtensionList(new Draw.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" })) { Embed = relationshipId, CompressionState = Draw.BlipCompressionValues.Print },
                                        new Draw.Stretch(new Draw.FillRectangle())),
                                    new Pic.ShapeProperties(
                                        new Draw.Transform2D(
                                            new Draw.Offset() { X = 0L, Y = 0L },
                                            new Draw.Extents() { Cx = emuWidth, Cy = emuHeight }),
                                        new Draw.PresetGeometry(new Draw.AdjustValueList()) { Preset = Draw.ShapeTypeValues.Rectangle }))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                    { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, EditId = "50D07946" });

            return element;
        }

        private static Wp.Inline ConvertAnchorToInline(Wp.Anchor anchor)
        {
            var inline = new Wp.Inline();
            var extent = anchor.Elements<Wp.Extent>().FirstOrDefault();
            if (extent != null) inline.AppendChild((Wp.Extent)extent.CloneNode(true));
            var effectExtent = anchor.Elements<Wp.EffectExtent>().FirstOrDefault();
            if (effectExtent != null) inline.AppendChild((Wp.EffectExtent)effectExtent.CloneNode(true));
            var docPr = anchor.Elements<Wp.DocProperties>().FirstOrDefault();
            if (docPr != null) inline.AppendChild((Wp.DocProperties)docPr.CloneNode(true));
            var cNvGraphicFramePr = anchor.Elements<Wp.NonVisualGraphicFrameDrawingProperties>().FirstOrDefault();
            if (cNvGraphicFramePr != null) inline.AppendChild((Wp.NonVisualGraphicFrameDrawingProperties)cNvGraphicFramePr.CloneNode(true));
            var graphic = anchor.Elements<Draw.Graphic>().FirstOrDefault();
            if (graphic != null) inline.AppendChild((Draw.Graphic)graphic.CloneNode(true));
            return inline;
        }

        private static Table CreateEquationTable(Paragraph originalEqPara)
        {
            Table table = new Table();
            TableProperties tblPr = new TableProperties(
                new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" },
                new TableBorders(
                    new TopBorder() { Val = BorderValues.None, Size = 0 },
                    new BottomBorder() { Val = BorderValues.None, Size = 0 },
                    new LeftBorder() { Val = BorderValues.None, Size = 0 },
                    new RightBorder() { Val = BorderValues.None, Size = 0 },
                    new InsideHorizontalBorder() { Val = BorderValues.None, Size = 0 },
                    new InsideVerticalBorder() { Val = BorderValues.None, Size = 0 }
                )
            );
            table.AppendChild(tblPr);

            TableGrid tblGrid = new TableGrid(
                new GridColumn() { Width = "750" },
                new GridColumn() { Width = "3500" },
                new GridColumn() { Width = "750" }
            );
            table.AppendChild(tblGrid);

            TableRow tr = new TableRow();
            TableCell tc1 = new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "750" }), new Paragraph());

            Paragraph clonedEq = (Paragraph)originalEqPara.CloneNode(true);
            var pPr = clonedEq.ParagraphProperties ?? clonedEq.AppendChild(new ParagraphProperties());
            var jc = pPr.Elements<Justification>().FirstOrDefault() ?? pPr.AppendChild(new Justification());
            jc.Val = JustificationValues.Center;
            TableCell tc2 = new TableCell(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "3500" }), clonedEq);

            Paragraph pSeq = new Paragraph(
                new ParagraphProperties(new Justification() { Val = JustificationValues.Right }),
                new Run(new Text("PT ")),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(@" SEQ PT \* ARABIC ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.End })
            );
            TableCell tc3 = new TableCell(
                new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "750" },
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                ), pSeq
            );

            tr.Append(tc1, tc2, tc3);
            table.AppendChild(tr);
            return table;
        }

        private static Paragraph CreateHeading(string text, string fontName)
        {
            return new Paragraph(new Run(new Text(text)) { RunProperties = new RunProperties(new RunFonts { Ascii = fontName, HighAnsi = fontName }, new Bold(), new FontSize() { Val = "32" }) })
            {
                ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center })
            };
        }

        private static Paragraph CreateFieldCode(string instruction)
        {
            return new Paragraph(
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(instruction) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.End })
            );
        }
    }
}