using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocStyleVerify.API.Models;
using DocStyleVerify.API.DTOs;
using System.Text.Json;
using System.Security.Cryptography;
using System.Text;

namespace DocStyleVerify.API.Services
{
    public partial class DocumentProcessingService
    {
        // Table helper methods
        private void ExtractTableBorders(TableBorders tableBorders, TextStyle textStyle)
        {
            var borderParts = new List<string>();

            if (tableBorders.TopBorder != null)
            {
                textStyle.TableBorderStyle = tableBorders.TopBorder.Val?.ToString() ?? "";
                textStyle.TableBorderColor = ConvertColorToHex(tableBorders.TopBorder.Color?.Value ?? "000000");
                textStyle.TableBorderWidth = ConvertEighthPointsToPoints(tableBorders.TopBorder.Size?.Value ?? 0);
                borderParts.Add("Top");
            }

            if (tableBorders.BottomBorder != null) borderParts.Add("Bottom");
            if (tableBorders.LeftBorder != null) borderParts.Add("Left");
            if (tableBorders.RightBorder != null) borderParts.Add("Right");
            if (tableBorders.InsideHorizontalBorder != null) borderParts.Add("InsideHorizontal");
            if (tableBorders.InsideVerticalBorder != null) borderParts.Add("InsideVertical");

            textStyle.BorderDirections = string.Join(",", borderParts);
        }

        private void ExtractTableCellBorders(TableCellBorders cellBorders, TextStyle textStyle)
        {
            var borderParts = new List<string>();

            if (cellBorders.TopBorder != null)
            {
                textStyle.BorderStyle = cellBorders.TopBorder.Val?.ToString() ?? "";
                textStyle.BorderColor = ConvertColorToHex(cellBorders.TopBorder.Color?.Value ?? "000000");
                textStyle.BorderWidth = ConvertEighthPointsToPoints(cellBorders.TopBorder.Size?.Value ?? 0);
                borderParts.Add("Top");
            }

            if (cellBorders.BottomBorder != null) borderParts.Add("Bottom");
            if (cellBorders.LeftBorder != null) borderParts.Add("Left");
            if (cellBorders.RightBorder != null) borderParts.Add("Right");

            textStyle.BorderDirections = string.Join(",", borderParts);
        }

        private void ExtractTableCellMargins(TableCellMarginDefault cellMargins, TextStyle textStyle)
        {
            if (cellMargins.TopMargin != null)
                textStyle.TableCellTopMargin = ConvertTwipsToPoints(ParseStringToInt(cellMargins.TopMargin.Width?.Value));
            
            if (cellMargins.BottomMargin != null)
                textStyle.TableCellBottomMargin = ConvertTwipsToPoints(ParseStringToInt(cellMargins.BottomMargin.Width?.Value));
            
            if (cellMargins.StartMargin != null)
                textStyle.TableCellLeftMargin = ConvertTwipsToPoints(ParseStringToInt(cellMargins.StartMargin.Width?.Value));
            
            if (cellMargins.EndMargin != null)
                textStyle.TableCellRightMargin = ConvertTwipsToPoints(ParseStringToInt(cellMargins.EndMargin.Width?.Value));
        }

        private void ExtractIndividualTableCellMargins(TableCellMargin cellMargin, TextStyle textStyle)
        {
            if (cellMargin.TopMargin != null)
                textStyle.TableCellTopMargin = ConvertTwipsToPoints(ParseStringToInt(cellMargin.TopMargin.Width?.Value));
            
            if (cellMargin.BottomMargin != null)
                textStyle.TableCellBottomMargin = ConvertTwipsToPoints(ParseStringToInt(cellMargin.BottomMargin.Width?.Value));
            
            if (cellMargin.StartMargin != null)
                textStyle.TableCellLeftMargin = ConvertTwipsToPoints(ParseStringToInt(cellMargin.StartMargin.Width?.Value));
            
            if (cellMargin.EndMargin != null)
                textStyle.TableCellRightMargin = ConvertTwipsToPoints(ParseStringToInt(cellMargin.EndMargin.Width?.Value));
        }

        private float ConvertTableWidthToPoints(TableWidth tableWidth)
        {
            if (tableWidth.Type?.Value == TableWidthUnitValues.Dxa)
            {
                return ConvertTwipsToPoints(int.Parse(tableWidth.Width?.Value ?? "0"));
            }
            else if (tableWidth.Type?.Value == TableWidthUnitValues.Pct)
            {
                return (int.Parse(tableWidth.Width?.Value ?? "0")) / 50f; // Convert percentage (50ths) to percentage
            }
            
            return int.Parse(tableWidth.Width?.Value ?? "0");
        }

        private float ConvertTableCellWidthToPoints(TableCellWidth cellWidth)
        {
            if (cellWidth.Type?.Value == TableWidthUnitValues.Dxa)
            {
                return ConvertTwipsToPoints(int.Parse(cellWidth.Width?.Value ?? "0"));
            }
            else if (cellWidth.Type?.Value == TableWidthUnitValues.Pct)
            {
                return (int.Parse(cellWidth.Width?.Value ?? "0")) / 50f;
            }
            
            return int.Parse(cellWidth.Width?.Value ?? "0");
        }

        // Context creation methods
        private FormattingContext CreateTableFormattingContext(Table table, int tableIndex)
        {
            return new FormattingContext
            {
                ElementType = "Table",
                TableIndex = tableIndex,
                StructuralRole = "Table",
                SampleText = GetTableSampleText(table),
                ContextKey = $"Table:{tableIndex}"
            };
        }

        private FormattingContext CreateTableRowFormattingContext(TableRow row, int tableIndex, int rowIndex)
        {
            return new FormattingContext
            {
                ElementType = "TableRow",
                TableIndex = tableIndex,
                RowIndex = rowIndex,
                StructuralRole = "TableRow",
                SampleText = GetTableRowSampleText(row),
                ContextKey = $"Table:{tableIndex}:Row:{rowIndex}"
            };
        }

        private FormattingContext CreateTableCellFormattingContext(TableCell cell, int tableIndex, int rowIndex, int cellIndex)
        {
            return new FormattingContext
            {
                ElementType = "TableCell",
                TableIndex = tableIndex,
                RowIndex = rowIndex,
                CellIndex = cellIndex,
                StructuralRole = "TableCell",
                SampleText = GetTableCellSampleText(cell),
                ContextKey = $"Table:{tableIndex}:Row:{rowIndex}:Cell:{cellIndex}"
            };
        }

        private FormattingContext CreateTableCellParagraphFormattingContext(Paragraph paragraph, int tableIndex, int rowIndex, int cellIndex, int paragraphIndex)
        {
            return new FormattingContext
            {
                ElementType = "TableCellParagraph",
                TableIndex = tableIndex,
                RowIndex = rowIndex,
                CellIndex = cellIndex,
                ParagraphIndex = paragraphIndex,
                StructuralRole = "TableCellParagraph",
                SampleText = GetSampleText(paragraph),
                ContextKey = $"Table:{tableIndex}:Row:{rowIndex}:Cell:{cellIndex}:Paragraph:{paragraphIndex}"
            };
        }

        private FormattingContext CreateHeaderFormattingContext(Paragraph paragraph, int headerIndex, int paragraphIndex, string headerType)
        {
            return new FormattingContext
            {
                ElementType = "Header",
                ParagraphIndex = paragraphIndex,
                StructuralRole = "Header",
                SampleText = GetSampleText(paragraph),
                HeaderFooterType = headerType,
                IsInHeaderFooter = true,
                DocumentPartType = "Header",
                ContextKey = $"Header:{headerIndex}:Paragraph:{paragraphIndex}"
            };
        }

        private FormattingContext CreateFooterFormattingContext(Paragraph paragraph, int footerIndex, int paragraphIndex, string footerType)
        {
            return new FormattingContext
            {
                ElementType = "Footer",
                ParagraphIndex = paragraphIndex,
                StructuralRole = "Footer",
                SampleText = GetSampleText(paragraph),
                HeaderFooterType = footerType,
                IsInHeaderFooter = true,
                DocumentPartType = "Footer",
                ContextKey = $"Footer:{footerIndex}:Paragraph:{paragraphIndex}"
            };
        }

        // Header/Footer helper methods
        private string DetermineHeaderType(HeaderPart headerPart)
        {
            // This would need to be determined based on the relationship or context
            // For now, return a generic type
            return "Header";
        }

        private string DetermineFooterType(FooterPart footerPart)
        {
            // This would need to be determined based on the relationship or context
            // For now, return a generic type
            return "Footer";
        }

        private async Task ExtractHeaderTableStylesAsync(Table table, int headerIndex, int tableIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            await ExtractTableStyleDetailsAsync(table, tableIndex, styles, patterns, templateId, createdBy);
            
            // Update the context for all table-related styles to indicate they're in a header
            var recentStyles = styles.TakeLast(10).ToList(); // Approximate number of styles just added
            foreach (var style in recentStyles)
            {
                if (style.FormattingContext != null)
                {
                    style.FormattingContext.IsInHeaderFooter = true;
                    style.FormattingContext.DocumentPartType = "Header";
                    style.FormattingContext.ContextKey = $"Header:{headerIndex}:{style.FormattingContext.ContextKey}";
                }
            }
        }

        private async Task ExtractFooterTableStylesAsync(Table table, int footerIndex, int tableIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            await ExtractTableStyleDetailsAsync(table, tableIndex, styles, patterns, templateId, createdBy);
            
            // Update the context for all table-related styles to indicate they're in a footer
            var recentStyles = styles.TakeLast(10).ToList(); // Approximate number of styles just added
            foreach (var style in recentStyles)
            {
                if (style.FormattingContext != null)
                {
                    style.FormattingContext.IsInHeaderFooter = true;
                    style.FormattingContext.DocumentPartType = "Footer";
                    style.FormattingContext.ContextKey = $"Footer:{footerIndex}:{style.FormattingContext.ContextKey}";
                }
            }
        }

        // Direct formatting helper methods
        private string GetRunText(Run run)
        {
            var text = run.InnerText;
            return text;
        }

        private DirectFormatPattern? ExtractRunDirectFormatting(Run run, int textStyleId, int paragraphIndex, int runIndex, string createdBy)
        {
            var runProperties = run.RunProperties;
            if (runProperties == null) return null;

            // Check if run has any direct formatting different from default
            var hasDirectFormatting = HasDirectFormatting(runProperties);
            if (!hasDirectFormatting) return null;

            var directPattern = new DirectFormatPattern
            {
                TextStyleId = textStyleId,
                PatternName = $"Run_{paragraphIndex}_{runIndex}",
                Context = $"Paragraph:{paragraphIndex},Run:{runIndex}",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                SampleText = GetRunText(run)
            };

            // Extract direct formatting properties
            ExtractDirectRunFormatting(runProperties, directPattern);

            // Count occurrences of similar patterns
            directPattern.OccurrenceCount = 1; // Will be updated during aggregation

            return directPattern;
        }

        private bool HasDirectFormatting(RunProperties runProperties)
        {
            // Check for visual formatting properties that represent direct formatting
            // Exclude style references and proofing properties as they are not direct formatting
            return runProperties.Bold != null ||
                   runProperties.Italic != null ||
                   runProperties.Underline != null ||
                   runProperties.Strike != null ||
                   runProperties.Color != null ||
                   runProperties.FontSize != null ||
                   runProperties.RunFonts != null ||
                   runProperties.Highlight != null ||
                   runProperties.Spacing != null ||
                   runProperties.Caps != null ||
                   runProperties.SmallCaps != null ||
                   runProperties.Shadow != null ||
                   runProperties.Outline != null ||
                   runProperties.Emboss != null ||
                   runProperties.Imprint != null ||
                   runProperties.DoubleStrike != null ||
                   runProperties.Position != null ||
                   runProperties.Kern != null ||
                   runProperties.VerticalTextAlignment != null ||
                   runProperties.RightToLeftText != null ||
                   runProperties.Languages != null ||
                   runProperties.CharacterScale != null;
            // Note: Removed Vanish and WebHidden as they are typically document metadata, not visual formatting
        }

        private void ExtractDirectRunFormatting(RunProperties runProperties, DirectFormatPattern directPattern)
        {
            // Font properties
            if (runProperties.RunFonts != null)
            {
                directPattern.FontFamily = runProperties.RunFonts.Ascii?.Value ?? 
                                          runProperties.RunFonts.HighAnsi?.Value ?? 
                                          runProperties.RunFonts.ComplexScript?.Value ?? 
                                          runProperties.RunFonts.EastAsia?.Value;
            }

            if (runProperties.FontSize != null)
            {
                var fontSizeValue = runProperties.FontSize.Val?.Value;
                if (!string.IsNullOrEmpty(fontSizeValue))
                {
                    directPattern.FontSize = float.Parse(fontSizeValue) / 2f; // Half-points to points
                }
            }

            // Boolean formatting properties with proper default handling
            if (runProperties.Bold != null)
                directPattern.IsBold = runProperties.Bold.Val?.Value ?? true;
            
            if (runProperties.Italic != null)
                directPattern.IsItalic = runProperties.Italic.Val?.Value ?? true;
            
            if (runProperties.Underline != null)
                directPattern.IsUnderline = true; // Underline presence indicates it's applied
            
            if (runProperties.Strike != null)
                directPattern.IsStrikethrough = runProperties.Strike.Val?.Value ?? true;
            
            if (runProperties.Caps != null)
                directPattern.IsAllCaps = runProperties.Caps.Val?.Value ?? true;
            
            if (runProperties.SmallCaps != null)
                directPattern.IsSmallCaps = runProperties.SmallCaps.Val?.Value ?? true;
            
            if (runProperties.Shadow != null)
                directPattern.HasShadow = runProperties.Shadow.Val?.Value ?? true;
            
            if (runProperties.Outline != null)
                directPattern.HasOutline = runProperties.Outline.Val?.Value ?? true;

            // Color properties
            if (runProperties.Color != null)
            {
                var colorValue = runProperties.Color.Val?.Value;
                if (!string.IsNullOrEmpty(colorValue) && colorValue != "auto")
                {
                    directPattern.Color = ConvertColorToHex(colorValue);
                }
            }

            if (runProperties.Highlight != null)
                directPattern.Highlighting = runProperties.Highlight.Val?.ToString();

            // Spacing and positioning
            if (runProperties.Spacing != null)
            {
                var spacingValue = runProperties.Spacing.Val?.Value;
                if (spacingValue.HasValue)
                {
                    directPattern.CharacterSpacing = ConvertTwipsToPoints(spacingValue.Value);
                }
            }

            if (runProperties.Languages != null)
                directPattern.Language = runProperties.Languages.Val?.Value;

            // Additional character formatting not previously captured
            if (runProperties.Position != null)
            {
                // Store character position (subscript/superscript) in a custom way
                var positionValue = runProperties.Position.Val?.Value;
                if (!string.IsNullOrEmpty(positionValue))
                {
                    var position = int.Parse(positionValue);
                    if (position != 0)
                    {
                        directPattern.SampleText += $" [Position: {ConvertTwipsToPoints(position)}pt]";
                    }
                }
            }

            if (runProperties.CharacterScale != null)
            {
                var scale = runProperties.CharacterScale.Val?.Value ?? 100;
                if (scale != 100)
                {
                    directPattern.SampleText += $" [Scale: {scale}%]";
                }
            }
        }

        // New method to extract paragraph-level direct formatting patterns
        private DirectFormatPattern? ExtractParagraphDirectFormatting(Paragraph paragraph, int textStyleId, int paragraphIndex, string createdBy)
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties == null) return null;

            // Check if paragraph has direct formatting beyond what's in the style
            var hasDirectFormatting = HasParagraphDirectFormatting(paragraphProperties);
            if (!hasDirectFormatting) return null;

            var directPattern = new DirectFormatPattern
            {
                TextStyleId = textStyleId,
                PatternName = $"Paragraph_{paragraphIndex}_Direct",
                Context = $"Paragraph:{paragraphIndex}",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                SampleText = GetSampleText(paragraph),
                OccurrenceCount = 1
            };

            // Extract paragraph-level direct formatting
            ExtractDirectParagraphFormatting(paragraphProperties, directPattern);

            return directPattern;
        }

        private bool HasParagraphDirectFormatting(ParagraphProperties paragraphProperties)
        {
            // Check for visual paragraph formatting that represents direct formatting
            // Don't count having a style reference as direct formatting
            return paragraphProperties.Justification != null ||
                   paragraphProperties.SpacingBetweenLines != null ||
                   paragraphProperties.Indentation != null ||
                   paragraphProperties.ParagraphBorders != null ||
                   paragraphProperties.Shading != null ||
                   paragraphProperties.Tabs != null ||
                   paragraphProperties.KeepNext != null ||
                   paragraphProperties.KeepLines != null ||
                   paragraphProperties.WidowControl != null ||
                   paragraphProperties.SuppressLineNumbers != null ||
                   paragraphProperties.PageBreakBefore != null;
        }

        private void ExtractDirectParagraphFormatting(ParagraphProperties paragraphProperties, DirectFormatPattern directPattern)
        {
            // Alignment
            if (paragraphProperties.Justification != null)
                directPattern.Alignment = paragraphProperties.Justification.Val?.ToString() ?? "Left";

            // Spacing
            if (paragraphProperties.SpacingBetweenLines != null)
            {
                var spacing = paragraphProperties.SpacingBetweenLines;
                if (spacing.Before != null)
                    directPattern.SpacingBefore = ConvertTwipsToPoints(int.Parse(spacing.Before.Value ?? "0"));
                if (spacing.After != null)
                    directPattern.SpacingAfter = ConvertTwipsToPoints(int.Parse(spacing.After.Value ?? "0"));
                if (spacing.Line != null)
                    directPattern.LineSpacing = int.Parse(spacing.Line.Value ?? "240") / 240f;
            }

            // Indentation
            if (paragraphProperties.Indentation != null)
            {
                var indent = paragraphProperties.Indentation;
                if (indent.Left != null)
                    directPattern.IndentationLeft = ConvertTwipsToPoints(int.Parse(indent.Left.Value ?? "0"));
                if (indent.Right != null)
                    directPattern.IndentationRight = ConvertTwipsToPoints(int.Parse(indent.Right.Value ?? "0"));
                if (indent.FirstLine != null)
                    directPattern.FirstLineIndent = ConvertTwipsToPoints(int.Parse(indent.FirstLine.Value ?? "0"));
            }
        }

        private async Task ExtractRunPatternsFromParagraph(Paragraph paragraph, int textStyleId, List<DirectFormatPattern> patterns, string createdBy, int contextIndex, int paragraphIndex)
        {
            var runs = paragraph.Descendants<Run>().ToList();
            for (int runIndex = 0; runIndex < runs.Count; runIndex++)
            {
                var run = runs[runIndex];
                var directPattern = ExtractRunDirectFormatting(run, textStyleId, paragraphIndex, runIndex, createdBy);
                if (directPattern != null)
                {
                    directPattern.Context = $"Context:{contextIndex},Paragraph:{paragraphIndex},Run:{runIndex}";
                    patterns.Add(directPattern);
                }
            }
        }

        // Sample text helper methods
        private string GetTableSampleText(Table table)
        {
            var text = table.InnerText;
            return text;
        }

        private string GetTableRowSampleText(TableRow row)
        {
            var text = row.InnerText;
            return text;
        }

        private string GetTableCellSampleText(TableCell cell)
        {
            var text = cell.InnerText;
            return text;
        }

        // List/Numbering implementation methods
        private TextStyle? ExtractNumberingStyleAsync(AbstractNum abstractNum, int numIndex, List<TextStyle> styles, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Numbering",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"Numbering_{numIndex}"
            };

            // Extract numbering properties
            ExtractAbstractNumberingProperties(abstractNum, textStyle);

            textStyle.FormattingContext = CreateNumberingFormattingContext(abstractNum, numIndex);
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);

            return textStyle;
        }

        private void ExtractAbstractNumberingProperties(AbstractNum abstractNum, TextStyle textStyle)
        {
            textStyle.AbstractNumberingId = abstractNum.AbstractNumberId?.Value ?? 0;

            // Extract all level definitions for comprehensive numbering information
            var levelDefinitions = new List<string>();
            foreach (var level in abstractNum.Descendants<Level>())
            {
                var levelInfo = ExtractLevelDefinition(level);
                if (!string.IsNullOrEmpty(levelInfo))
                {
                    levelDefinitions.Add(levelInfo);
                }
            }

            if (levelDefinitions.Any())
            {
                textStyle.ListBulletChar = string.Join(" | ", levelDefinitions);
            }

            // Extract level properties (focusing on level 0 for primary properties)
            var primaryLevel = abstractNum.Descendants<Level>().FirstOrDefault();
            if (primaryLevel != null)
            {
                textStyle.ListLevel = primaryLevel.LevelIndex?.Value ?? 0;
                
                if (primaryLevel.NumberingFormat != null)
                    textStyle.ListNumberFormat = primaryLevel.NumberingFormat.Val?.ToString() ?? "";

                if (primaryLevel.LevelText != null)
                    textStyle.ListLevelText = primaryLevel.LevelText.Val?.Value ?? "";

                if (primaryLevel.StartNumberingValue != null)
                    textStyle.ListStartValue = (primaryLevel.StartNumberingValue.Val?.Value ?? 1).ToString();

                if (primaryLevel.LevelJustification != null)
                    textStyle.Alignment = primaryLevel.LevelJustification.Val?.ToString() ?? "Left";

                // Extract level paragraph properties
                if (primaryLevel.PreviousParagraphProperties != null)
                {
                    var pPr = primaryLevel.PreviousParagraphProperties;
                    if (pPr.Indentation != null)
                    {
                        textStyle.ListIndentPosition = ConvertTwipsToPoints(ParseStringToInt(pPr.Indentation.Left?.Value));
                        textStyle.ListTabPosition = ConvertTwipsToPoints(ParseStringToInt(pPr.Indentation.Hanging?.Value));
                    }
                }

                // Extract level run properties for bullet character
                if (primaryLevel.NumberingSymbolRunProperties != null)
                {
                    var rPr = primaryLevel.NumberingSymbolRunProperties;
                    if (rPr.RunFonts != null)
                        textStyle.ListBulletFont = rPr.RunFonts.Ascii?.Value ?? "";
                }
            }
        }

        private string ExtractLevelDefinition(Level level)
        {
            var levelIndex = level.LevelIndex?.Value ?? 0;
            var numberFormat = level.NumberingFormat?.Val?.ToString() ?? "bullet";
            var levelText = level.LevelText?.Val?.Value ?? "";
            var startValue = level.StartNumberingValue?.Val?.Value ?? 1;
            
            // Get level indentation
            var leftIndent = 0f;
            var hangingIndent = 0f;
            
            if (level.PreviousParagraphProperties?.Indentation != null)
            {
                var indent = level.PreviousParagraphProperties.Indentation;
                if (indent.Left != null)
                    leftIndent = ConvertTwipsToPoints(ParseStringToInt(indent.Left.Value));
                if (indent.Hanging != null)
                    hangingIndent = ConvertTwipsToPoints(ParseStringToInt(indent.Hanging.Value));
            }

            // Build comprehensive level definition
            var levelDef = $"L{levelIndex}:{numberFormat}";
            
            if (!string.IsNullOrEmpty(levelText))
                levelDef += $":'{levelText}'";
                
            if (startValue != 1)
                levelDef += $":Start{startValue}";
                
            if (leftIndent > 0 || hangingIndent > 0)
                levelDef += $":Indent({leftIndent}pt,{hangingIndent}pt)";

            return levelDef;
        }

        private TextStyle? ExtractListParagraphStyleAsync(Paragraph paragraph, int listIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var paragraphStyle = ExtractParagraphStyleAsync(paragraph, templateId, createdBy, listIndex).Result;
            if (paragraphStyle != null)
            {
                paragraphStyle.StyleType = "ListItem";
                
                // Extract comprehensive list-specific properties
                var numPr = paragraph.ParagraphProperties?.NumberingProperties;
                if (numPr != null)
                {
                    ExtractNumberingProperties(numPr, paragraphStyle);
                    
                    // Get the actual list level and numbering ID for better naming
                    var listLevel = numPr.NumberingLevelReference?.Val?.Value ?? 0;
                    var numberingId = numPr.NumberingId?.Val?.Value ?? 0;
                    
                    paragraphStyle.Name = $"ListItem_L{listLevel}_N{numberingId}_{listIndex}";
                    
                    // Add more detailed list information to the sample text
                    var listText = GetListItemText(paragraph, listLevel, numberingId);
                    if (!string.IsNullOrEmpty(listText))
                    {
                        paragraphStyle.FormattingContext.SampleText = listText;
                    }
                }
                else
                {
                    paragraphStyle.Name = $"ListItem_{listIndex}";
                }

                paragraphStyle.FormattingContext = CreateListFormattingContext(paragraph, listIndex);
                
                // Use temporary ID for pattern linking
                var tempStyleId = -(styles.Count + 1);
                paragraphStyle.Id = tempStyleId;
                styles.Add(paragraphStyle);
                
                // Extract direct formatting patterns from runs in list paragraphs
                var runIndex = 0;
                foreach (var run in paragraph.Descendants<Run>())
                {
                    var directPattern = ExtractRunDirectFormatting(run, tempStyleId, listIndex, runIndex, createdBy);
                    if (directPattern != null)
                    {
                        // Update context to indicate this is within a list item with level info
                        var listLevel = numPr?.NumberingLevelReference?.Val?.Value ?? 0;
                        var numberingId = numPr?.NumberingId?.Val?.Value ?? 0;
                        directPattern.Context = $"ListItem:L{listLevel}:N{numberingId}:P{listIndex}:R{runIndex}";
                        patterns.Add(directPattern);
                    }
                    runIndex++;
                }
                
                // Reset ID for EF
                paragraphStyle.Id = 0;
            }
            return paragraphStyle;
        }

        private string GetListItemText(Paragraph paragraph, int listLevel, int numberingId)
        {
            var text = paragraph.InnerText;
            var prefix = $"[L{listLevel}:N{numberingId}] ";
            
            if (string.IsNullOrEmpty(text))
                return prefix + "[Empty list item]";
                
            return prefix + text;
        }

        private void ExtractNumberingProperties(NumberingProperties numPr, TextStyle textStyle)
        {
            if (numPr.NumberingLevelReference != null)
                textStyle.ListLevel = numPr.NumberingLevelReference.Val?.Value ?? 0;

            if (numPr.NumberingId != null)
                textStyle.NumberingId = numPr.NumberingId.Val?.Value ?? 0;

            // Try to get the actual numbering format from the document
            try
            {
                var numberingId = textStyle.NumberingId;
                var listLevel = textStyle.ListLevel;
                
                if (numberingId > 0)
                {
                    var numberingFormat = GetNumberingFormat(numberingId, listLevel);
                    if (!string.IsNullOrEmpty(numberingFormat))
                    {
                        textStyle.ListBulletChar = numberingFormat;
                    }
                }
            }
            catch (Exception ex)
            {
                // Log but don't fail the extraction
                System.Diagnostics.Debug.WriteLine($"Failed to extract numbering format: {ex.Message}");
            }
        }

        private string GetNumberingFormat(int numberingId, int listLevel)
        {
            // This would need access to the document's numbering part
            // For now, return a placeholder that indicates the numbering structure
            return $"NumId:{numberingId}:Level:{listLevel}";
        }

        private FormattingContext CreateNumberingFormattingContext(AbstractNum abstractNum, int numIndex)
        {
            return new FormattingContext
            {
                ElementType = "Numbering",
                StructuralRole = "Numbering",
                SampleText = $"Numbering definition {numIndex}",
                ContextKey = $"Numbering:{numIndex}"
            };
        }

        private FormattingContext CreateListFormattingContext(Paragraph paragraph, int listIndex)
        {
            // Extract list level information for better context
            var listLevel = 0;
            var numberingId = 0;
            
            var numPr = paragraph.ParagraphProperties?.NumberingProperties;
            if (numPr != null)
            {
                listLevel = numPr.NumberingLevelReference?.Val?.Value ?? 0;
                numberingId = numPr.NumberingId?.Val?.Value ?? 0;
            }
            
            return new FormattingContext
            {
                ElementType = "ListItem",
                ParagraphIndex = listIndex,
                StructuralRole = "ListItem",
                SampleText = GetSampleText(paragraph),
                ContextKey = $"ListItem:{listIndex}:Level:{listLevel}:NumId:{numberingId}"
            };
        }

        // Image/Drawing implementation methods
        private TextStyle? ExtractDrawingStyleAsync(Drawing drawing, int drawingIndex, List<TextStyle> styles, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Drawing",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"Drawing_{drawingIndex}"
            };

            // Extract drawing properties
            var inline = drawing.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>();
            var anchor = drawing.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor>();

            if (inline != null)
            {
                ExtractInlineDrawingProperties(inline, textStyle);
            }
            else if (anchor != null)
            {
                ExtractAnchorDrawingProperties(anchor, textStyle);
            }

            textStyle.FormattingContext = CreateDrawingFormattingContext(drawing, drawingIndex);
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private TextStyle? ExtractPictureStyleAsync(DocumentFormat.OpenXml.Drawing.Pictures.Picture picture, int pictureIndex, List<TextStyle> styles, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Picture",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"Picture_{pictureIndex}"
            };

            // Extract picture properties
            ExtractPictureProperties(picture, textStyle);

            textStyle.FormattingContext = CreatePictureFormattingContext(picture, pictureIndex);
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private void ExtractInlineDrawingProperties(DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline inline, TextStyle textStyle)
        {
            if (inline.Extent != null)
            {
                textStyle.ImageWidth = ConvertEmuToPoints(inline.Extent.Cx?.Value ?? 0);
                textStyle.ImageHeight = ConvertEmuToPoints(inline.Extent.Cy?.Value ?? 0);
            }

            textStyle.ImageWrapType = "Inline";
            textStyle.ImagePosition = "Inline";

            if (inline.DocProperties != null)
            {
                textStyle.ImageAltText = inline.DocProperties.Description?.Value ?? "";
                textStyle.ImageTitle = inline.DocProperties.Title?.Value ?? "";
            }
        }

        private void ExtractAnchorDrawingProperties(DocumentFormat.OpenXml.Drawing.Wordprocessing.Anchor anchor, TextStyle textStyle)
        {
            if (anchor.Extent != null)
            {
                textStyle.ImageWidth = ConvertEmuToPoints(anchor.Extent.Cx?.Value ?? 0);
                textStyle.ImageHeight = ConvertEmuToPoints(anchor.Extent.Cy?.Value ?? 0);
            }

            textStyle.ImageWrapType = "Anchored";
            
            // Extract positioning information
            var positionH = anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalPosition>();
            var positionV = anchor.GetFirstChild<DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalPosition>();
            
            var positionInfo = new List<string>();
            if (positionH != null) positionInfo.Add($"H:{positionH.RelativeFrom}");
            if (positionV != null) positionInfo.Add($"V:{positionV.RelativeFrom}");
            textStyle.ImagePosition = string.Join(",", positionInfo);

            // Note: anchor.DocProperties doesn't exist, getting properties from different location
            textStyle.ImageAltText = "";
            textStyle.ImageTitle = "";
        }

        private void ExtractPictureProperties(DocumentFormat.OpenXml.Drawing.Pictures.Picture picture, TextStyle textStyle)
        {
            // Extract basic picture properties
            var blipFill = picture.BlipFill;
            if (blipFill?.Blip != null)
            {
                var blip = blipFill.Blip;
                // Picture-specific properties would be extracted here
            }
        }

        private float ConvertEmuToPoints(long emu)
        {
            // 1 point = 12700 EMUs (English Metric Units)
            return emu / 12700f;
        }

        private FormattingContext CreateDrawingFormattingContext(Drawing drawing, int drawingIndex)
        {
            return new FormattingContext
            {
                ElementType = "Drawing",
                StructuralRole = "Drawing",
                SampleText = $"Drawing {drawingIndex}",
                ContextKey = $"Drawing:{drawingIndex}"
            };
        }

        private FormattingContext CreatePictureFormattingContext(DocumentFormat.OpenXml.Drawing.Pictures.Picture picture, int pictureIndex)
        {
            return new FormattingContext
            {
                ElementType = "Picture",
                StructuralRole = "Picture",
                SampleText = $"Picture {pictureIndex}",
                ContextKey = $"Picture:{pictureIndex}"
            };
        }

        // Section implementation methods
        private TextStyle? ExtractSectionPropertiesStyleAsync(SectionProperties sectionProperties, int sectionIndex, List<TextStyle> styles, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Section",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"Section_{sectionIndex}"
            };

            // Extract page size
            var pageSize = sectionProperties.GetFirstChild<PageSize>();
            if (pageSize != null)
            {
                textStyle.PageWidth = ConvertTwipsToPoints(ParseUIntToInt(pageSize.Width?.Value) == 0 ? 12240 : ParseUIntToInt(pageSize.Width?.Value)); // Default 8.5"
                textStyle.PageHeight = ConvertTwipsToPoints(ParseUIntToInt(pageSize.Height?.Value) == 0 ? 15840 : ParseUIntToInt(pageSize.Height?.Value)); // Default 11"
                textStyle.Orientation = pageSize.Orient?.ToString() ?? "Portrait";
            }

            // Extract page margins
            var pageMargin = sectionProperties.GetFirstChild<PageMargin>();
            if (pageMargin != null)
            {
                textStyle.TopMargin = ConvertTwipsToPoints(pageMargin.Top?.Value ?? 1440);
                textStyle.BottomMargin = ConvertTwipsToPoints(pageMargin.Bottom?.Value ?? 1440);
                textStyle.LeftMargin = ConvertTwipsToPoints(ParseUIntToInt(pageMargin.Left?.Value) == 0 ? 1440 : ParseUIntToInt(pageMargin.Left?.Value));
                textStyle.RightMargin = ConvertTwipsToPoints(ParseUIntToInt(pageMargin.Right?.Value) == 0 ? 1440 : ParseUIntToInt(pageMargin.Right?.Value));
                textStyle.HeaderDistance = ConvertTwipsToPoints(ParseUIntToInt(pageMargin.Header?.Value) == 0 ? 720 : ParseUIntToInt(pageMargin.Header?.Value));
                textStyle.FooterDistance = ConvertTwipsToPoints(ParseUIntToInt(pageMargin.Footer?.Value) == 0 ? 720 : ParseUIntToInt(pageMargin.Footer?.Value));
                // Note: MirrorMargins property doesn't exist on PageMargin, using default
                textStyle.MirrorMargins = false;
            }

            // Extract columns
            var columns = sectionProperties.GetFirstChild<Columns>();
            if (columns != null)
            {
                textStyle.ColumnCount = columns.ColumnCount?.Value ?? 1;
                textStyle.ColumnSpacing = ConvertTwipsToPoints(ParseStringToInt(columns.Space?.Value));
                textStyle.HasColumnSeparator = columns.Separator?.Value ?? false;
            }

            // Extract section type
            var sectionType = sectionProperties.GetFirstChild<SectionType>();
            if (sectionType != null)
            {
                textStyle.SectionType = sectionType.Val?.ToString() ?? "NextPage";
            }

            // Extract page numbering
            var pageNumberType = sectionProperties.GetFirstChild<PageNumberType>();
            if (pageNumberType != null)
            {
                textStyle.PageNumberFormat = pageNumberType.Format?.ToString() ?? "Decimal";
                textStyle.PageNumberStart = pageNumberType.Start?.Value ?? 1;
            }

            textStyle.FormattingContext = CreateSectionFormattingContext(sectionProperties, sectionIndex);
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private FormattingContext CreateSectionFormattingContext(SectionProperties sectionProperties, int sectionIndex)
        {
            return new FormattingContext
            {
                ElementType = "Section",
                SectionIndex = sectionIndex,
                StructuralRole = "Section",
                SampleText = $"Section {sectionIndex}",
                ContextKey = $"Section:{sectionIndex}"
            };
        }

        // Field implementation methods
        private TextStyle? ExtractSimpleFieldStyleAsync(SimpleField field, int fieldIndex, List<TextStyle> styles, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Field",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"SimpleField_{fieldIndex}"
            };

            textStyle.FieldType = "Simple";
            textStyle.FieldCode = field.Instruction?.Value ?? "";
            textStyle.FieldResult = field.InnerText;
            textStyle.FieldLocked = false; // field.Lock doesn't exist, using default
            textStyle.FieldDirty = field.Dirty?.Value ?? false;

            textStyle.FormattingContext = CreateFieldFormattingContext(field, fieldIndex, "Simple");
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private TextStyle? ExtractComplexFieldStyleAsync(FieldChar fieldChar, int fieldIndex, List<TextStyle> styles, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Field",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"ComplexField_{fieldIndex}"
            };

            textStyle.FieldType = "Complex";
            textStyle.FieldLocked = fieldChar.FieldLock?.Value ?? false;
            textStyle.FieldDirty = fieldChar.Dirty?.Value ?? false;

            textStyle.FormattingContext = CreateFieldFormattingContext(fieldChar, fieldIndex, "Complex");
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private FormattingContext CreateFieldFormattingContext(OpenXmlElement field, int fieldIndex, string fieldType)
        {
            return new FormattingContext
            {
                ElementType = "Field",
                StructuralRole = "Field",
                SampleText = $"{fieldType} Field {fieldIndex}",
                ContextKey = $"Field:{fieldIndex}"
            };
        }

        // Hyperlink implementation methods
        private TextStyle? ExtractHyperlinkStyleAsync(Hyperlink hyperlink, int hyperlinkIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Hyperlink",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"Hyperlink_{hyperlinkIndex}"
            };

            // Extract hyperlink properties
            textStyle.HyperlinkTarget = hyperlink.Id?.Value ?? "";
            textStyle.HyperlinkTooltip = hyperlink.Tooltip?.Value ?? "";
            textStyle.HyperlinkVisited = hyperlink.History?.Value ?? false;

            // Extract run formatting from hyperlink content
            var run = hyperlink.GetFirstChild<Run>();
            if (run?.RunProperties != null)
            {
                ExtractRunPropertiesIntoStyle(run.RunProperties, textStyle);
            }

            textStyle.FormattingContext = CreateHyperlinkFormattingContext(hyperlink, hyperlinkIndex);
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private FormattingContext CreateHyperlinkFormattingContext(Hyperlink hyperlink, int hyperlinkIndex)
        {
            return new FormattingContext
            {
                ElementType = "Hyperlink",
                StructuralRole = "Hyperlink",
                SampleText = hyperlink.InnerText,
                ContextKey = $"Hyperlink:{hyperlinkIndex}"
            };
        }

        // Document Style Definitions implementation methods
        private TextStyle? ExtractDocumentStyleAsync(Style style, int styleIndex, List<TextStyle> styles, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = $"DocumentStyle_{style.Type?.ToString() ?? "Unknown"}",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = style.StyleName?.Val?.Value ?? $"Style_{styleIndex}"
            };

            // Extract style properties
            if (style.StyleParagraphProperties != null)
            {
                ExtractStyleParagraphProperties(style.StyleParagraphProperties, textStyle);
            }

            if (style.StyleRunProperties != null)
            {
                ExtractStyleRunProperties(style.StyleRunProperties, textStyle);
            }

            // Extract style metadata
            textStyle.BasedOnStyle = style.BasedOn?.Val?.Value ?? "";
            textStyle.Priority = style.UIPriority?.Val?.Value ?? 0;

            textStyle.FormattingContext = CreateDocumentStyleFormattingContext(style, styleIndex);
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private void ExtractStyleParagraphProperties(StyleParagraphProperties stylePPr, TextStyle textStyle)
        {
            if (stylePPr.Justification != null)
                textStyle.Alignment = stylePPr.Justification.Val?.ToString() ?? "Left";

            if (stylePPr.SpacingBetweenLines != null)
            {
                if (stylePPr.SpacingBetweenLines.Before != null)
                    textStyle.SpacingBefore = ConvertTwipsToPoints(int.Parse(stylePPr.SpacingBetweenLines.Before.Value ?? "0"));
                if (stylePPr.SpacingBetweenLines.After != null)
                    textStyle.SpacingAfter = ConvertTwipsToPoints(int.Parse(stylePPr.SpacingBetweenLines.After.Value ?? "0"));
                if (stylePPr.SpacingBetweenLines.Line != null)
                    textStyle.LineSpacing = int.Parse(stylePPr.SpacingBetweenLines.Line.Value ?? "240") / 240f;
            }

            if (stylePPr.Indentation != null)
            {
                if (stylePPr.Indentation.Left != null)
                    textStyle.IndentationLeft = ConvertTwipsToPoints(int.Parse(stylePPr.Indentation.Left.Value ?? "0"));
                if (stylePPr.Indentation.Right != null)
                    textStyle.IndentationRight = ConvertTwipsToPoints(int.Parse(stylePPr.Indentation.Right.Value ?? "0"));
                if (stylePPr.Indentation.FirstLine != null)
                    textStyle.FirstLineIndent = ConvertTwipsToPoints(int.Parse(stylePPr.Indentation.FirstLine.Value ?? "0"));
            }
        }

        private void ExtractStyleRunProperties(StyleRunProperties styleRPr, TextStyle textStyle)
        {
            if (styleRPr.RunFonts != null)
                textStyle.FontFamily = styleRPr.RunFonts.Ascii?.Value ?? "Calibri";

            if (styleRPr.FontSize != null)
                textStyle.FontSize = float.Parse(styleRPr.FontSize.Val?.Value ?? "22") / 2f;

            textStyle.IsBold = styleRPr.Bold != null && (styleRPr.Bold.Val?.Value ?? true);
            textStyle.IsItalic = styleRPr.Italic != null && (styleRPr.Italic.Val?.Value ?? true);
            textStyle.IsUnderline = styleRPr.Underline != null;

            if (styleRPr.Color != null)
                textStyle.Color = ConvertColorToHex(styleRPr.Color.Val?.Value ?? "000000");
        }

        private FormattingContext CreateDocumentStyleFormattingContext(Style style, int styleIndex)
        {
            return new FormattingContext
            {
                ElementType = "DocumentStyle",
                StructuralRole = "DocumentStyle",
                SampleText = style.StyleName?.Val?.Value ?? $"Style {styleIndex}",
                ContextKey = $"DocumentStyle:{styleIndex}"
            };
        }

        // Content Control implementation methods - simplified
        private TextStyle? ExtractBlockContentControlStyleAsync(SdtBlock contentControl, int ccIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "ContentControl",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"ContentControl_{ccIndex}"
            };

            textStyle.FormattingContext = CreateContentControlFormattingContext(contentControl, ccIndex, "Block");
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private TextStyle? ExtractInlineContentControlStyleAsync(SdtRun contentControl, int ccIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "InlineContentControl",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"InlineContentControl_{ccIndex}"
            };

            textStyle.FormattingContext = CreateContentControlFormattingContext(contentControl, ccIndex, "Inline");
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private TextStyle? ExtractCellContentControlStyleAsync(SdtCell contentControl, int ccIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "CellContentControl",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"CellContentControl_{ccIndex}"
            };

            textStyle.FormattingContext = CreateContentControlFormattingContext(contentControl, ccIndex, "Cell");
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private TextStyle? ExtractRowContentControlStyleAsync(SdtRow contentControl, int ccIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "RowContentControl",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"RowContentControl_{ccIndex}"
            };

            textStyle.FormattingContext = CreateContentControlFormattingContext(contentControl, ccIndex, "Row");
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            styles.Add(textStyle);
            return textStyle;
        }

        private FormattingContext CreateContentControlFormattingContext(OpenXmlElement contentControl, int ccIndex, string controlType)
        {
            return new FormattingContext
            {
                ElementType = "ContentControl",
                StructuralRole = "ContentControl",
                SampleText = GetContentControlSampleText(contentControl),
                ContentControlType = controlType,
                ContextKey = $"ContentControl:{ccIndex}"
            };
        }

        private FormattingContext CreateContentControlParagraphFormattingContext(Paragraph paragraph, int ccIndex, int paragraphIndex)
        {
            return new FormattingContext
            {
                ElementType = "ContentControlParagraph",
                ParagraphIndex = paragraphIndex,
                StructuralRole = "ContentControlParagraph",
                SampleText = GetSampleText(paragraph),
                ContentControlType = "Block",
                ContextKey = $"ContentControl:{ccIndex}:Paragraph:{paragraphIndex}"
            };
        }

        private string GetContentControlSampleText(OpenXmlElement contentControl)
        {
            var text = contentControl.InnerText;
            return text;
        }

        // Helper method to safely parse string values to int
        private int ParseStringToInt(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return 0;
            
            if (int.TryParse(value, out int result))
                return result;
            
            return 0;
        }

        // Helper method to safely parse uint values to int
        private int ParseUIntToInt(uint? value)
        {
            return (int)(value ?? 0);
        }



    }
} 