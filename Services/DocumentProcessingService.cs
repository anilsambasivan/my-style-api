using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocStyleVerify.API.Models;
using DocStyleVerify.API.DTOs;
using DocStyleVerify.API.Data;
using Microsoft.EntityFrameworkCore;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace DocStyleVerify.API.Services
{
    public partial class DocumentProcessingService : IDocumentProcessingService
    {
        private readonly DocStyleVerifyDbContext _context;
        private readonly ILogger<DocumentProcessingService> _logger;

        public DocumentProcessingService(DocStyleVerifyDbContext context, ILogger<DocumentProcessingService> logger)
        {
            _context = context;
            _logger = logger;
        }

        public async Task<TemplateExtractionResult> ExtractTemplateStylesAsync(Stream documentStream, string fileName, int templateId, string createdBy)
        {
            var result = new TemplateExtractionResult
            {
                TemplateId = templateId,
                TemplateName = fileName
            };

            try
            {
                // Extract styles without saving
                var (extractedStyles, extractedPatterns) = await ExtractStylesOnlyAsync(documentStream, templateId, createdBy);

                // Save to database only for template extraction
                await SaveExtractedStylesAsync(extractedStyles, extractedPatterns);

                result.TotalTextStyles = extractedStyles.Count;
                result.TotalDirectFormatPatterns = extractedPatterns.Count;
                result.ExtractedStyleTypes = extractedStyles.Select(s => s.StyleType).Distinct().ToList();
                result.Success = true;
                result.ProcessingMessages = new List<string> 
                { 
                    $"Processed {extractedStyles.Count} text styles",
                    $"Found {extractedPatterns.Count} direct format patterns",
                    $"Covered style types: {string.Join(", ", result.ExtractedStyleTypes)}"
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting template styles from {FileName}", fileName);
                result.ErrorMessage = ex.Message;
                result.Success = false;
            }

            return result;
        }

        /// <summary>
        /// Extracts styles from a document without saving to database. Used for both template extraction and verification.
        /// </summary>
        /// <param name="documentStream">Document stream to extract from</param>
        /// <param name="templateId">Template ID (use 0 for verification scenarios)</param>
        /// <param name="createdBy">User who initiated the extraction</param>
        /// <returns>Tuple containing extracted styles and patterns</returns>
        public async Task<(List<TextStyle> styles, List<DirectFormatPattern> patterns)> ExtractStylesOnlyAsync(Stream documentStream, int templateId, string createdBy)
        {
            var extractedStyles = new List<TextStyle>();
            var extractedPatterns = new List<DirectFormatPattern>();

            using var document = WordprocessingDocument.Open(documentStream, false);
            var mainPart = document.MainDocumentPart;
            
            if (mainPart == null)
                return (extractedStyles, extractedPatterns);

            // Extract all style types - same methods but without database saving
            await ExtractParagraphStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            await ExtractTableStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            await ExtractHeaderFooterStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            await ExtractContentControlStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            await ExtractListStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            ExtractImageStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            ExtractSectionStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            ExtractFieldStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            ExtractHyperlinkStylesAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);
            ExtractStyleDefinitionsAsync(mainPart, extractedStyles, extractedPatterns, templateId, createdBy);

            return (extractedStyles, extractedPatterns);
        }

        private async Task ExtractParagraphStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var document = mainPart.Document;
            var body = document.Body;
            if (body == null) return;

            var paragraphs = body.Descendants<Paragraph>().ToList();
            int paragraphIndex = 0;

            foreach (var paragraph in paragraphs)
            {
                var textStyle = await ExtractParagraphStyleAsync(paragraph, templateId, createdBy, paragraphIndex);
                if (textStyle != null)
                {
                    textStyle.FormattingContext = CreateFormattingContext(paragraph, paragraphIndex, "Paragraph");
                    styles.Add(textStyle);

                    // Extract direct formatting patterns from paragraph and runs
                    // Use a temporary reference (negative index) to link patterns to styles before saving
                    var tempStyleId = -(styles.Count + 1); // Use negative numbers as temporary IDs
                    textStyle.Id = tempStyleId; // Temporarily set this for pattern linking
                    
                    // Extract paragraph-level direct formatting
                    var paragraphDirectPattern = ExtractParagraphDirectFormatting(paragraph, tempStyleId, paragraphIndex, createdBy);
                    if (paragraphDirectPattern != null)
                    {
                        patterns.Add(paragraphDirectPattern);
                    }
                    
                    // Extract run-level direct formatting
                    var runIndex = 0;
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        var directPattern = ExtractRunDirectFormatting(run, tempStyleId, paragraphIndex, runIndex, createdBy);
                        if (directPattern != null)
                        {
                            patterns.Add(directPattern);
                        }
                        runIndex++;
                    }
                    
                    // Reset the ID to 0 for EF to generate the real ID
                    textStyle.Id = 0;
                }
                paragraphIndex++;
            }
        }

        private async Task<TextStyle?> ExtractParagraphStyleAsync(Paragraph paragraph, int templateId, string createdBy, int paragraphIndex)
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            
            // Create a text style even if paragraph properties are null - this is critical for capturing run-level direct formatting
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Paragraph",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow
            };

            // Check if paragraph has any runs with direct formatting
            var hasRunFormatting = paragraph.Descendants<Run>().Any(r => r.RunProperties != null);
            
            // If no paragraph properties and no run formatting, skip this paragraph
            if (paragraphProperties == null && !hasRunFormatting) 
                return null;

            // Check if paragraph references a named style and inherit properties
            string styleName = "Custom";
            string styleType = "Paragraph";
            
            if (paragraphProperties?.ParagraphStyleId != null)
            {
                var styleId = paragraphProperties.ParagraphStyleId.Val?.Value;
                if (!string.IsNullOrEmpty(styleId))
                {
                    styleName = GetStyleNameFromId(styleId);
                    styleType = DetermineStyleTypeFromName(styleName);
                    textStyle.BasedOnStyle = styleId;
                    
                    // Inherit properties from the document style definition
                    await InheritFromDocumentStyle(textStyle, styleId, paragraph);
                }
            }

            // Extract paragraph style properties (these override inherited properties)
            if (paragraphProperties != null)
            {
                ExtractParagraphPropertiesIntoStyle(paragraphProperties, textStyle);
            }

            // Extract run properties for character formatting (these override inherited properties)
            var firstRun = paragraph.Descendants<Run>().FirstOrDefault();
            if (firstRun?.RunProperties != null)
            {
                ExtractRunPropertiesIntoStyle(firstRun.RunProperties, textStyle);
            }
            else if (paragraphProperties == null && hasRunFormatting)
            {
                // For paragraphs without paragraph properties but with run formatting,
                // ensure we have some base formatting from document defaults
                if (string.IsNullOrEmpty(textStyle.FontFamily))
                {
                    textStyle.FontFamily = "Calibri"; // Word default
                }
                if (textStyle.FontSize == 0)
                {
                    textStyle.FontSize = 11; // Word default
                }
                if (string.IsNullOrEmpty(textStyle.Color))
                {
                    textStyle.Color = "000000"; // Black
                }
            }

            // Set proper style type and name
            textStyle.StyleType = styleType;
            
            // Handle naming for paragraphs with and without paragraph properties
            if (string.IsNullOrEmpty(textStyle.BasedOnStyle))
            {
                // For paragraphs without style references, check if they have run formatting
                if (hasRunFormatting && paragraphProperties == null)
                {
                    textStyle.Name = $"DirectFormat_Paragraph_{paragraphIndex}";
                    textStyle.StyleType = "DirectFormatting";
                }
                else
                {
                    textStyle.Name = $"Paragraph_{paragraphIndex}_{DateTime.UtcNow.Ticks % 10000}";
                }
            }
            else
            {
                textStyle.Name = $"{styleName}_{paragraphIndex}";
            }

            // Generate style signature
            textStyle.StyleSignature = GenerateStyleSignature(textStyle);

            return textStyle;
        }

        private string GetStyleNameFromId(string styleId)
        {
            // Common Word style ID to name mappings
            var styleMapping = new Dictionary<string, string>
            {
                { "Normal", "Normal" },
                { "Heading1", "Heading 1" },
                { "Heading2", "Heading 2" },
                { "Heading3", "Heading 3" },
                { "Heading4", "Heading 4" },
                { "Heading5", "Heading 5" },
                { "Heading6", "Heading 6" },
                { "Heading7", "Heading 7" },
                { "Heading8", "Heading 8" },
                { "Heading9", "Heading 9" },
                { "Title", "Title" },
                { "Subtitle", "Subtitle" },
                { "Quote", "Quote" },
                { "IntenseQuote", "Intense Quote" },
                { "ListParagraph", "List Paragraph" },
                { "Caption", "Caption" },
                { "TOCHeading", "TOC Heading" },
                { "TOC1", "TOC 1" },
                { "TOC2", "TOC 2" },
                { "TOC3", "TOC 3" },
                { "FootnoteText", "Footnote Text" },
                { "EndnoteText", "Endnote Text" },
                { "Header", "Header" },
                { "Footer", "Footer" }
            };

            return styleMapping.ContainsKey(styleId) ? styleMapping[styleId] : styleId;
        }

        private string DetermineStyleTypeFromName(string styleName)
        {
            if (styleName.StartsWith("Heading") || styleName.Contains("Title") || styleName.Contains("Subtitle"))
                return "Heading";
            
            if (styleName.Contains("List") || styleName.Contains("Bullet") || styleName.Contains("Number"))
                return "List";
            
            if (styleName.Contains("TOC") || styleName.Contains("Table of Contents"))
                return "TOC";
            
            if (styleName.Contains("Caption") || styleName.Contains("Figure"))
                return "Caption";
            
            if (styleName.Contains("Quote") || styleName.Contains("Blockquote"))
                return "Quote";
            
            if (styleName.Contains("Header") || styleName.Contains("Footer"))
                return "HeaderFooter";
            
            if (styleName.Contains("Footnote") || styleName.Contains("Endnote"))
                return "Note";
            
            return "Paragraph";
        }

        private async Task InheritFromDocumentStyle(TextStyle textStyle, string styleId, Paragraph paragraph)
        {
            try
            {
                // Get the document part that contains this paragraph
                var document = paragraph.Ancestors<Document>().FirstOrDefault();
                if (document == null) return;

                var mainPart = document.MainDocumentPart;
                if (mainPart?.StyleDefinitionsPart?.Styles == null) return;

                // Find the style definition
                var styleDefinition = mainPart.StyleDefinitionsPart.Styles
                    .Descendants<Style>()
                    .FirstOrDefault(s => s.StyleId?.Value == styleId);

                if (styleDefinition == null) return;

                // Inherit paragraph properties from style definition
                if (styleDefinition.StyleParagraphProperties != null)
                {
                    InheritParagraphPropertiesFromStyle(styleDefinition.StyleParagraphProperties, textStyle);
                }

                // Inherit run properties from style definition
                if (styleDefinition.StyleRunProperties != null)
                {
                    InheritRunPropertiesFromStyle(styleDefinition.StyleRunProperties, textStyle);
                }

                // Handle style inheritance chain (basedOn)
                if (styleDefinition.BasedOn?.Val?.Value != null)
                {
                    await InheritFromDocumentStyle(textStyle, styleDefinition.BasedOn.Val.Value, paragraph);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to inherit style properties for style {StyleId}", styleId);
            }
        }

        private void InheritParagraphPropertiesFromStyle(StyleParagraphProperties stylePPr, TextStyle textStyle)
        {
            // Only inherit if property is not already set (allows for override)
            if (string.IsNullOrEmpty(textStyle.Alignment) && stylePPr.Justification != null)
                textStyle.Alignment = stylePPr.Justification.Val?.ToString() ?? "Left";

            if (textStyle.SpacingBefore == 0 && stylePPr.SpacingBetweenLines?.Before != null)
                textStyle.SpacingBefore = ConvertTwipsToPoints(int.Parse(stylePPr.SpacingBetweenLines.Before.Value ?? "0"));

            if (textStyle.SpacingAfter == 0 && stylePPr.SpacingBetweenLines?.After != null)
                textStyle.SpacingAfter = ConvertTwipsToPoints(int.Parse(stylePPr.SpacingBetweenLines.After.Value ?? "0"));

            if (textStyle.LineSpacing == 0 && stylePPr.SpacingBetweenLines?.Line != null)
                textStyle.LineSpacing = int.Parse(stylePPr.SpacingBetweenLines.Line.Value ?? "240") / 240f;

            if (textStyle.IndentationLeft == 0 && stylePPr.Indentation?.Left != null)
                textStyle.IndentationLeft = ConvertTwipsToPoints(int.Parse(stylePPr.Indentation.Left.Value ?? "0"));

            if (textStyle.IndentationRight == 0 && stylePPr.Indentation?.Right != null)
                textStyle.IndentationRight = ConvertTwipsToPoints(int.Parse(stylePPr.Indentation.Right.Value ?? "0"));

            if (textStyle.FirstLineIndent == 0 && stylePPr.Indentation?.FirstLine != null)
                textStyle.FirstLineIndent = ConvertTwipsToPoints(int.Parse(stylePPr.Indentation.FirstLine.Value ?? "0"));

            // Inherit other properties
            if (!textStyle.KeepWithNext && stylePPr.KeepNext != null)
                textStyle.KeepWithNext = stylePPr.KeepNext.Val?.Value ?? false;

            if (!textStyle.KeepTogether && stylePPr.KeepLines != null)
                textStyle.KeepTogether = stylePPr.KeepLines.Val?.Value ?? false;

            if (!textStyle.WidowOrphanControl && stylePPr.WidowControl != null)
                textStyle.WidowOrphanControl = stylePPr.WidowControl.Val?.Value ?? false;
        }

        private void InheritRunPropertiesFromStyle(StyleRunProperties styleRPr, TextStyle textStyle)
        {
            // Only inherit if property is not already set (allows for override)
            if (string.IsNullOrEmpty(textStyle.FontFamily) && styleRPr.RunFonts != null)
            {
                var fontName = styleRPr.RunFonts.Ascii?.Value ?? 
                              styleRPr.RunFonts.HighAnsi?.Value ?? 
                              styleRPr.RunFonts.ComplexScript?.Value ?? 
                              styleRPr.RunFonts.EastAsia?.Value;
                
                if (!string.IsNullOrEmpty(fontName))
                {
                    textStyle.FontFamily = ResolveThemeFont(fontName);
                }
            }

            if (textStyle.FontSize == 0 && styleRPr.FontSize != null)
            {
                var fontSize = styleRPr.FontSize.Val?.Value;
                if (!string.IsNullOrEmpty(fontSize))
                {
                    textStyle.FontSize = float.Parse(fontSize) / 2f; // Half-points to points
                }
            }

            // Boolean properties - inherit if not explicitly set
            if (!textStyle.IsBold && styleRPr.Bold != null)
                textStyle.IsBold = styleRPr.Bold.Val?.Value ?? true;

            if (!textStyle.IsItalic && styleRPr.Italic != null)
                textStyle.IsItalic = styleRPr.Italic.Val?.Value ?? true;

            if (!textStyle.IsUnderline && styleRPr.Underline != null)
                textStyle.IsUnderline = true;

            if (!textStyle.IsStrikethrough && styleRPr.Strike != null)
                textStyle.IsStrikethrough = styleRPr.Strike.Val?.Value ?? true;

            if (!textStyle.IsAllCaps && styleRPr.Caps != null)
                textStyle.IsAllCaps = styleRPr.Caps.Val?.Value ?? true;

            if (!textStyle.IsSmallCaps && styleRPr.SmallCaps != null)
                textStyle.IsSmallCaps = styleRPr.SmallCaps.Val?.Value ?? true;

            if (!textStyle.HasShadow && styleRPr.Shadow != null)
                textStyle.HasShadow = styleRPr.Shadow.Val?.Value ?? true;

            if (!textStyle.HasOutline && styleRPr.Outline != null)
                textStyle.HasOutline = styleRPr.Outline.Val?.Value ?? true;

            // Color - inherit if not already set
            if (string.IsNullOrEmpty(textStyle.Color) && styleRPr.Color != null)
            {
                var colorValue = styleRPr.Color.Val?.Value;
                if (!string.IsNullOrEmpty(colorValue) && colorValue != "auto")
                {
                    textStyle.Color = ConvertColorToHex(colorValue);
                }
            }

            // Note: StyleRunProperties doesn't have Highlight property like RunProperties
            // Highlighting would be handled through RunProperties, not StyleRunProperties
        }

        private void ExtractParagraphPropertiesIntoStyle(ParagraphProperties paragraphProperties, TextStyle textStyle)
        {
            // Alignment
            if (paragraphProperties.Justification != null)
            {
                textStyle.Alignment = paragraphProperties.Justification.Val?.ToString() ?? "Left";
            }

            // Spacing
            if (paragraphProperties.SpacingBetweenLines != null)
            {
                if (paragraphProperties.SpacingBetweenLines.Before != null)
                    textStyle.SpacingBefore = ConvertTwipsToPoints(int.Parse(paragraphProperties.SpacingBetweenLines.Before.Value ?? "0"));
                if (paragraphProperties.SpacingBetweenLines.After != null)
                    textStyle.SpacingAfter = ConvertTwipsToPoints(int.Parse(paragraphProperties.SpacingBetweenLines.After.Value ?? "0"));
                if (paragraphProperties.SpacingBetweenLines.Line != null)
                    textStyle.LineSpacing = int.Parse(paragraphProperties.SpacingBetweenLines.Line.Value ?? "240") / 240f;
            }

            // Indentation
            if (paragraphProperties.Indentation != null)
            {
                if (paragraphProperties.Indentation.Left != null)
                    textStyle.IndentationLeft = ConvertTwipsToPoints(int.Parse(paragraphProperties.Indentation.Left.Value ?? "0"));
                if (paragraphProperties.Indentation.Right != null)
                    textStyle.IndentationRight = ConvertTwipsToPoints(int.Parse(paragraphProperties.Indentation.Right.Value ?? "0"));
                if (paragraphProperties.Indentation.FirstLine != null)
                    textStyle.FirstLineIndent = ConvertTwipsToPoints(int.Parse(paragraphProperties.Indentation.FirstLine.Value ?? "0"));
            }

            // Tabs
            if (paragraphProperties.Tabs != null)
            {
                textStyle.TabStops = ExtractTabStops(paragraphProperties.Tabs);
            }

            // Paragraph control properties
            if (paragraphProperties.WidowControl != null)
                textStyle.WidowOrphanControl = paragraphProperties.WidowControl.Val?.Value ?? false;
            if (paragraphProperties.KeepNext != null)
                textStyle.KeepWithNext = paragraphProperties.KeepNext.Val?.Value ?? false;
            if (paragraphProperties.KeepLines != null)
                textStyle.KeepTogether = paragraphProperties.KeepLines.Val?.Value ?? false;

            // Borders
            if (paragraphProperties.ParagraphBorders != null)
            {
                ExtractParagraphBorders(paragraphProperties.ParagraphBorders, textStyle);
            }
        }

        private void ExtractRunPropertiesIntoStyle(RunProperties runProperties, TextStyle textStyle)
        {
            // Font - capture actual font with theme resolution
            if (runProperties.RunFonts != null)
            {
                var fontName = runProperties.RunFonts.Ascii?.Value ?? 
                              runProperties.RunFonts.HighAnsi?.Value ?? 
                              runProperties.RunFonts.ComplexScript?.Value ?? 
                              runProperties.RunFonts.EastAsia?.Value ?? 
                              "";
                
                // Resolve theme fonts
                if (!string.IsNullOrEmpty(fontName))
                {
                    textStyle.FontFamily = ResolveThemeFont(fontName);
                }
            }

            // Font size - capture actual size, don't fallback to defaults
            if (runProperties.FontSize != null)
            {
                var fontSize = runProperties.FontSize.Val?.Value;
                if (!string.IsNullOrEmpty(fontSize))
                {
                    textStyle.FontSize = float.Parse(fontSize) / 2f; // Half-points to points
                }
            }

            // Boolean properties - only set if explicitly defined
            textStyle.IsBold = runProperties.Bold != null && (runProperties.Bold.Val?.Value ?? true);
            textStyle.IsItalic = runProperties.Italic != null && (runProperties.Italic.Val?.Value ?? true);
            textStyle.IsUnderline = runProperties.Underline != null;
            textStyle.IsStrikethrough = runProperties.Strike != null && (runProperties.Strike.Val?.Value ?? true);
            textStyle.IsAllCaps = runProperties.Caps != null && (runProperties.Caps.Val?.Value ?? true);
            textStyle.IsSmallCaps = runProperties.SmallCaps != null && (runProperties.SmallCaps.Val?.Value ?? true);
            textStyle.HasShadow = runProperties.Shadow != null && (runProperties.Shadow.Val?.Value ?? true);
            textStyle.HasOutline = runProperties.Outline != null && (runProperties.Outline.Val?.Value ?? true);

            // Color - capture actual colors with theme resolution
            if (runProperties.Color != null)
            {
                var colorValue = runProperties.Color.Val?.Value;
                if (!string.IsNullOrEmpty(colorValue) && colorValue != "auto")
                {
                    textStyle.Color = ConvertColorToHex(colorValue);
                }
            }

            // Highlighting
            if (runProperties.Highlight != null)
            {
                textStyle.Highlighting = runProperties.Highlight.Val?.ToString() ?? "";
            }

            // Character spacing
            if (runProperties.Spacing != null)
            {
                textStyle.CharacterSpacing = ConvertTwipsToPoints(runProperties.Spacing.Val?.Value ?? 0);
            }

            // Language information
            if (runProperties.Languages != null)
            {
                textStyle.Language = runProperties.Languages.Val?.Value ?? "";
            }
        }

        private List<Models.TabStop> ExtractTabStops(Tabs tabs)
        {
            var tabStops = new List<Models.TabStop>();
            
            foreach (var tab in tabs.Descendants<DocumentFormat.OpenXml.Wordprocessing.TabStop>())
            {
                tabStops.Add(new Models.TabStop
                {
                    Position = ConvertTwipsToPoints(tab.Position?.Value ?? 0),
                    Alignment = tab.Val?.ToString() ?? "Left",
                    Leader = tab.Leader?.ToString() ?? "None"
                });
            }

            return tabStops;
        }

        private void ExtractParagraphBorders(ParagraphBorders borders, TextStyle textStyle)
        {
            var borderParts = new List<string>();

            if (borders.TopBorder != null)
            {
                textStyle.BorderStyle = borders.TopBorder.Val?.ToString() ?? "";
                textStyle.BorderColor = ConvertColorToHex(borders.TopBorder.Color?.Value ?? "000000");
                textStyle.BorderWidth = ConvertEighthPointsToPoints(borders.TopBorder.Size?.Value ?? 0);
                borderParts.Add("Top");
            }

            if (borders.BottomBorder != null)
            {
                borderParts.Add("Bottom");
            }

            if (borders.LeftBorder != null)
            {
                borderParts.Add("Left");
            }

            if (borders.RightBorder != null)
            {
                borderParts.Add("Right");
            }

            textStyle.BorderDirections = string.Join(",", borderParts);
        }

        // Utility methods
        protected float ConvertTwipsToPoints(int twips)
        {
            return twips / 20f; // 1 point = 20 twips
        }

        protected float ConvertEighthPointsToPoints(uint eighthPoints)
        {
            return eighthPoints / 8f;
        }

        public string ConvertColorToHex(string colorValue)
        {
            if (string.IsNullOrEmpty(colorValue) || colorValue.Equals("auto", StringComparison.OrdinalIgnoreCase))
            {
                return "000000"; // Default to black
            }

            // Handle theme colors
            var themeColorValue = ResolveThemeColor(colorValue);
            if (!string.IsNullOrEmpty(themeColorValue))
            {
                return themeColorValue;
            }

            // Handle hex colors
            if (colorValue.Length == 6 && colorValue.All(c => "0123456789ABCDEFabcdef".Contains(c)))
            {
                return colorValue.ToUpper();
            }

            return "000000"; // Default to black for unrecognized colors
        }

        private string ResolveThemeColor(string themeColorValue)
        {
            // Common theme color mappings (Office default theme colors)
            var themeColors = new Dictionary<string, string>
            {
                { "accent1", "4F81BD" },     // Blue
                { "accent2", "F79646" },     // Orange  
                { "accent3", "9BBB59" },     // Green
                { "accent4", "8064A2" },     // Purple
                { "accent5", "4BACC6" },     // Aqua
                { "accent6", "F24C4C" },     // Red
                { "background1", "FFFFFF" },  // White
                { "background2", "F2F2F2" },  // Light Gray
                { "text1", "000000" },       // Black
                { "text2", "1F497D" },       // Dark Blue
                { "dark1", "000000" },       // Black
                { "dark2", "1F497D" },       // Dark Blue
                { "light1", "FFFFFF" },      // White
                { "light2", "EEECE1" },      // Beige
                { "hyperlink", "0000FF" },   // Blue
                { "followedHyperlink", "800080" } // Purple
            };

            return themeColors.ContainsKey(themeColorValue.ToLower()) ? 
                   themeColors[themeColorValue.ToLower()] : 
                   "";
        }

        protected string ResolveThemeFont(string themeFontValue)
        {
            // Common theme font mappings (Office default theme fonts)
            var themeFonts = new Dictionary<string, string>
            {
                { "+mn-lt", "Calibri" },      // Minor Latin theme font
                { "+mj-lt", "Cambria" },      // Major Latin theme font
                { "+mn-ea", "Calibri" },      // Minor East Asian theme font
                { "+mj-ea", "Cambria" },      // Major East Asian theme font
                { "+mn-cs", "Calibri" },      // Minor Complex Script theme font
                { "+mj-cs", "Cambria" },      // Major Complex Script theme font
                { "minorFont", "Calibri" },   // Minor theme font
                { "majorFont", "Cambria" }    // Major theme font
            };

            return themeFonts.ContainsKey(themeFontValue.ToLower()) ? 
                   themeFonts[themeFontValue.ToLower()] : 
                   themeFontValue; // Return original if not a theme font
        }



        private FormattingContext CreateFormattingContext(Paragraph paragraph, int paragraphIndex, string elementType)
        {
            var context = new FormattingContext
            {
                ElementType = elementType,
                ParagraphIndex = paragraphIndex,
                SectionIndex = 0, // Will be updated based on section
                StructuralRole = DetermineStructuralRole(paragraph),
                SampleText = GetSampleText(paragraph)
            };

            context.ContextKey = GenerateContextKey(context);
            return context;
        }

        private string DetermineStructuralRole(Paragraph paragraph)
        {
            var paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties?.ParagraphStyleId != null)
            {
                var styleId = paragraphProperties.ParagraphStyleId.Val?.Value ?? "";
                
                if (styleId.Contains("Heading") || styleId.Contains("Title"))
                    return "Heading";
                if (styleId.Contains("TOC"))
                    return "TableOfContents";
                if (styleId.Contains("Caption"))
                    return "Caption";
                if (styleId.Contains("Quote"))
                    return "Quote";
            }

            return "Body";
        }

        protected string GetSampleText(Paragraph paragraph)
        {
            var text = paragraph.InnerText;
            return text;
        }

        private string GenerateContextKey(FormattingContext context)
        {
            var keyParts = new List<string>();
            
            if (context.SectionIndex >= 0)
                keyParts.Add($"Section:{context.SectionIndex}");
            
            if (context.TableIndex >= 0)
            {
                keyParts.Add($"Table:{context.TableIndex}");
                if (context.RowIndex >= 0)
                {
                    keyParts.Add($"Row:{context.RowIndex}");
                    if (context.CellIndex >= 0)
                        keyParts.Add($"Cell:{context.CellIndex}");
                }
            }
            
            if (context.ParagraphIndex >= 0)
                keyParts.Add($"Paragraph:{context.ParagraphIndex}");
            
            if (context.RunIndex >= 0)
                keyParts.Add($"Run:{context.RunIndex}");

            return string.Join(":", keyParts);
        }

        public string GenerateStyleSignature(TextStyle style)
        {
            var signatureData = new
            {
                style.FontFamily,
                style.FontSize,
                style.IsBold,
                style.IsItalic,
                style.IsUnderline,
                style.Color,
                style.Alignment,
                style.SpacingBefore,
                style.SpacingAfter,
                style.IndentationLeft,
                style.IndentationRight,
                style.FirstLineIndent,
                style.LineSpacing,
                style.StyleType
            };

            var json = JsonSerializer.Serialize(signatureData);
            using var sha256 = SHA256.Create();
            var hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(json));
            return Convert.ToBase64String(hashBytes)[..32]; // First 32 chars
        }

        // COMPREHENSIVE TABLE STYLES IMPLEMENTATION
        private async Task ExtractTableStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var tables = mainPart.Document.Body?.Descendants<Table>().ToList() ?? new List<Table>();
            
            for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
            {
                var table = tables[tableIndex];
                await ExtractTableStyleDetailsAsync(table, tableIndex, styles, patterns, templateId, createdBy);
            }
        }

        private async Task ExtractTableStyleDetailsAsync(Table table, int tableIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            // Extract table-level properties
            var tableProperties = table.GetFirstChild<TableProperties>();
            if (tableProperties != null)
            {
                var tableStyle = await CreateTableStyleAsync(tableProperties, templateId, createdBy, tableIndex);
                if (tableStyle != null)
                {
                    tableStyle.FormattingContext = CreateTableFormattingContext(table, tableIndex);
                    styles.Add(tableStyle);
                }
            }

            // Extract row styles
            var rows = table.Descendants<TableRow>().ToList();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var row = rows[rowIndex];
                await ExtractTableRowStyleAsync(row, tableIndex, rowIndex, styles, patterns, templateId, createdBy);
            }
        }

        private async Task<TextStyle?> CreateTableStyleAsync(TableProperties tableProperties, int templateId, string createdBy, int tableIndex)
        {
            var textStyle = new TextStyle
            {
                TemplateId = templateId,
                StyleType = "Table",
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Name = $"Table_{tableIndex}"
            };

            // Extract table style reference
            if (tableProperties.TableStyle != null)
            {
                textStyle.BasedOnStyle = tableProperties.TableStyle.Val?.Value ?? "";
                textStyle.Name = $"Table_{textStyle.BasedOnStyle}_{tableIndex}";
            }

            // Extract table borders
            if (tableProperties.TableBorders != null)
            {
                ExtractTableBorders(tableProperties.TableBorders, textStyle);
            }

            // Extract table width
            if (tableProperties.TableWidth != null)
            {
                textStyle.TableWidth = ConvertTableWidthToPoints(tableProperties.TableWidth);
            }

            // Extract table alignment
            if (tableProperties.TableJustification != null)
            {
                textStyle.TableAlignment = tableProperties.TableJustification.Val?.ToString() ?? "Left";
            }

            // Extract table spacing
            if (tableProperties.TableCellSpacing != null)
            {
                var spacingWidth = tableProperties.TableCellSpacing.Width?.Value;
                var spacingValue = spacingWidth?.ToString() ?? "0";
                textStyle.TableCellSpacingValue = ConvertTwipsToPoints(int.Parse(spacingValue));
            }

            // Extract table margins
            if (tableProperties.TableCellMarginDefault != null)
            {
                ExtractTableCellMargins(tableProperties.TableCellMarginDefault, textStyle);
            }

            // Extract table shading
            if (tableProperties.Shading != null)
            {
                textStyle.TableShadingColor = ConvertColorToHex(tableProperties.Shading.Fill?.Value ?? "FFFFFF");
                textStyle.TableShadingPattern = tableProperties.Shading.Val?.ToString() ?? "Clear";
            }

            // Extract table look (formatting options)
            if (tableProperties.TableLook != null)
            {
                var look = tableProperties.TableLook;
                var lookOptions = new List<string>();
                if (look.FirstRow?.Value == true) lookOptions.Add("FirstRow");
                if (look.LastRow?.Value == true) lookOptions.Add("LastRow");
                if (look.FirstColumn?.Value == true) lookOptions.Add("FirstColumn");
                if (look.LastColumn?.Value == true) lookOptions.Add("LastColumn");
                if (look.NoHorizontalBand?.Value == true) lookOptions.Add("NoHBand");
                if (look.NoVerticalBand?.Value == true) lookOptions.Add("NoVBand");
                
                // Store table look info in formatting context
                var contextSampleText = $"TableLook: {string.Join(", ", lookOptions)}";
            }

            // Extract table indent
            if (tableProperties.TableIndentation != null)
            {
                var indentWidth = tableProperties.TableIndentation.Width?.Value;
                var indentValue = indentWidth?.ToString() ?? "0";
                textStyle.IndentationLeft = ConvertTwipsToPoints(int.Parse(indentValue));
            }

            // Extract table positioning (floating tables)
            if (tableProperties.TablePositionProperties != null)
            {
                var pos = tableProperties.TablePositionProperties;
                textStyle.ImagePosition = $"HorzAnchor:{pos.HorizontalAnchor?.Value}, VertAnchor:{pos.VerticalAnchor?.Value}";
                if (pos.TablePositionXAlignment != null)
                    textStyle.Alignment = pos.TablePositionXAlignment.ToString() ?? "Left";
            }

            textStyle.StyleSignature = GenerateStyleSignature(textStyle);
            return textStyle;
        }

        private async Task ExtractTableRowStyleAsync(TableRow row, int tableIndex, int rowIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var rowProperties = row.TableRowProperties;
            if (rowProperties != null)
            {
                var rowStyle = new TextStyle
                {
                    TemplateId = templateId,
                    StyleType = "TableRow",
                    CreatedBy = createdBy,
                    CreatedOn = DateTime.UtcNow,
                    Name = $"TableRow_{tableIndex}_{rowIndex}"
                };

                // Extract row height
                var tableRowHeight = rowProperties.GetFirstChild<TableRowHeight>();
                if (tableRowHeight != null)
                {
                    var heightValue = tableRowHeight.Val?.Value;
                    if (heightValue.HasValue)
                    {
                        rowStyle.TableRowHeight = ConvertTwipsToPoints((int)heightValue.Value);
                    }
                    rowStyle.TableRowHeightRule = tableRowHeight.HeightType?.ToString() ?? "Auto";
                }

                // Extract row header
                var tableHeader = rowProperties.GetFirstChild<TableHeader>();
                rowStyle.IsTableHeader = tableHeader != null;

                // Extract row repeat on new page
                rowStyle.RepeatOnNewPage = tableHeader != null;

                rowStyle.FormattingContext = CreateTableRowFormattingContext(row, tableIndex, rowIndex);
                rowStyle.StyleSignature = GenerateStyleSignature(rowStyle);
                styles.Add(rowStyle);
            }

            // Extract cell styles
            var cells = row.Descendants<TableCell>().ToList();
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++)
            {
                var cell = cells[cellIndex];
                await ExtractTableCellStyleAsync(cell, tableIndex, rowIndex, cellIndex, styles, patterns, templateId, createdBy);
            }
        }

        private async Task ExtractTableCellStyleAsync(TableCell cell, int tableIndex, int rowIndex, int cellIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var cellProperties = cell.TableCellProperties;
            if (cellProperties != null)
            {
                var cellStyle = new TextStyle
                {
                    TemplateId = templateId,
                    StyleType = "TableCell",
                    CreatedBy = createdBy,
                    CreatedOn = DateTime.UtcNow,
                    Name = $"TableCell_{tableIndex}_{rowIndex}_{cellIndex}"
                };

                // Extract cell borders
                if (cellProperties.TableCellBorders != null)
                {
                    ExtractTableCellBorders(cellProperties.TableCellBorders, cellStyle);
                }

                // Extract cell shading
                if (cellProperties.Shading != null)
                {
                    cellStyle.TableShadingColor = ConvertColorToHex(cellProperties.Shading.Fill?.Value ?? "FFFFFF");
                    cellStyle.TableShadingPattern = cellProperties.Shading.Val?.ToString() ?? "Clear";
                }

                // Extract cell margins
                if (cellProperties.TableCellMargin != null)
                {
                    ExtractIndividualTableCellMargins(cellProperties.TableCellMargin, cellStyle);
                }

                // Extract cell width
                if (cellProperties.TableCellWidth != null)
                {
                    cellStyle.TableCellWidth = ConvertTableCellWidthToPoints(cellProperties.TableCellWidth);
                }

                // Extract vertical alignment
                if (cellProperties.TableCellVerticalAlignment != null)
                {
                    cellStyle.VerticalAlignment = cellProperties.TableCellVerticalAlignment.Val?.ToString() ?? "Top";
                }

                // Extract text direction
                if (cellProperties.TextDirection != null)
                {
                    cellStyle.TextDirection = cellProperties.TextDirection.Val?.ToString() ?? "LeftToRight";
                }

                cellStyle.FormattingContext = CreateTableCellFormattingContext(cell, tableIndex, rowIndex, cellIndex);
                cellStyle.StyleSignature = GenerateStyleSignature(cellStyle);
                styles.Add(cellStyle);
            }

            // Extract paragraph styles within the cell
            var paragraphs = cell.Descendants<Paragraph>().ToList();
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++)
            {
                var paragraph = paragraphs[paragraphIndex];
                var paragraphStyle = await ExtractParagraphStyleAsync(paragraph, templateId, createdBy, paragraphIndex);
                if (paragraphStyle != null)
                {
                    paragraphStyle.FormattingContext = CreateTableCellParagraphFormattingContext(paragraph, tableIndex, rowIndex, cellIndex, paragraphIndex);
                    paragraphStyle.StyleType = "TableCellParagraph";
                    paragraphStyle.Name = $"TableCellParagraph_{tableIndex}_{rowIndex}_{cellIndex}_{paragraphIndex}";
                    styles.Add(paragraphStyle);
                }
            }
        }

        // COMPREHENSIVE HEADER/FOOTER STYLES IMPLEMENTATION
        private async Task ExtractHeaderFooterStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            // Extract header styles
            var headerParts = mainPart.HeaderParts?.ToList() ?? new List<HeaderPart>();
            for (int headerIndex = 0; headerIndex < headerParts.Count; headerIndex++)
            {
                var headerPart = headerParts[headerIndex];
                await ExtractHeaderStylesAsync(headerPart, headerIndex, styles, patterns, templateId, createdBy);
            }

            // Extract footer styles
            var footerParts = mainPart.FooterParts?.ToList() ?? new List<FooterPart>();
            for (int footerIndex = 0; footerIndex < footerParts.Count; footerIndex++)
            {
                var footerPart = footerParts[footerIndex];
                await ExtractFooterStylesAsync(footerPart, footerIndex, styles, patterns, templateId, createdBy);
            }
        }

        private async Task ExtractHeaderStylesAsync(HeaderPart headerPart, int headerIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var header = headerPart.Header;
            if (header == null) return;

            var paragraphs = header.Descendants<Paragraph>().ToList();
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++)
            {
                var paragraph = paragraphs[paragraphIndex];
                var textStyle = await ExtractParagraphStyleAsync(paragraph, templateId, createdBy, paragraphIndex);
                if (textStyle != null)
                {
                    textStyle.StyleType = "Header";
                    textStyle.Name = $"Header_{headerIndex}_{paragraphIndex}";
                    textStyle.FormattingContext = CreateHeaderFormattingContext(paragraph, headerIndex, paragraphIndex, DetermineHeaderType(headerPart));
                    
                    // Use temporary ID for pattern linking
                    var tempStyleId = -(styles.Count + 1);
                    textStyle.Id = tempStyleId;
                    styles.Add(textStyle);

                    // Extract direct formatting patterns from runs
                    await ExtractRunPatternsFromParagraph(paragraph, tempStyleId, patterns, createdBy, headerIndex, paragraphIndex);
                    
                    // Reset ID for EF
                    textStyle.Id = 0;
                }
            }

            // Extract tables in headers
            var tables = header.Descendants<Table>().ToList();
            for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
            {
                var table = tables[tableIndex];
                await ExtractHeaderTableStylesAsync(table, headerIndex, tableIndex, styles, patterns, templateId, createdBy);
            }
        }

        private async Task ExtractFooterStylesAsync(FooterPart footerPart, int footerIndex, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var footer = footerPart.Footer;
            if (footer == null) return;

            var paragraphs = footer.Descendants<Paragraph>().ToList();
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++)
            {
                var paragraph = paragraphs[paragraphIndex];
                var textStyle = await ExtractParagraphStyleAsync(paragraph, templateId, createdBy, paragraphIndex);
                if (textStyle != null)
                {
                    textStyle.StyleType = "Footer";
                    textStyle.Name = $"Footer_{footerIndex}_{paragraphIndex}";
                    textStyle.FormattingContext = CreateFooterFormattingContext(paragraph, footerIndex, paragraphIndex, DetermineFooterType(footerPart));
                    
                    // Use temporary ID for pattern linking
                    var tempStyleId = -(styles.Count + 1);
                    textStyle.Id = tempStyleId;
                    styles.Add(textStyle);

                    // Extract direct formatting patterns from runs
                    await ExtractRunPatternsFromParagraph(paragraph, tempStyleId, patterns, createdBy, footerIndex, paragraphIndex);
                    
                    // Reset ID for EF
                    textStyle.Id = 0;
                }
            }

            // Extract tables in footers
            var tables = footer.Descendants<Table>().ToList();
            for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
            {
                var table = tables[tableIndex];
                await ExtractFooterTableStylesAsync(table, footerIndex, tableIndex, styles, patterns, templateId, createdBy);
            }
        }

        // COMPREHENSIVE CONTENT CONTROL STYLES IMPLEMENTATION
        private async Task ExtractContentControlStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            // Extract block-level content controls
            var blockContentControls = mainPart.Document.Body?.Descendants<SdtBlock>().ToList() ?? new List<SdtBlock>();
            for (int ccIndex = 0; ccIndex < blockContentControls.Count; ccIndex++)
            {
                var contentControl = blockContentControls[ccIndex];
                ExtractBlockContentControlStyleAsync(contentControl, ccIndex, styles, patterns, templateId, createdBy);
            }

            // Extract inline content controls
            var inlineContentControls = mainPart.Document.Body?.Descendants<SdtRun>().ToList() ?? new List<SdtRun>();
            for (int ccIndex = 0; ccIndex < inlineContentControls.Count; ccIndex++)
            {
                var contentControl = inlineContentControls[ccIndex];
                ExtractInlineContentControlStyleAsync(contentControl, ccIndex, styles, patterns, templateId, createdBy);
            }

            // Extract cell-level content controls
            var cellContentControls = mainPart.Document.Body?.Descendants<SdtCell>().ToList() ?? new List<SdtCell>();
            for (int ccIndex = 0; ccIndex < cellContentControls.Count; ccIndex++)
            {
                var contentControl = cellContentControls[ccIndex];
                ExtractCellContentControlStyleAsync(contentControl, ccIndex, styles, patterns, templateId, createdBy);
            }

            // Extract row-level content controls
            var rowContentControls = mainPart.Document.Body?.Descendants<SdtRow>().ToList() ?? new List<SdtRow>();
            for (int ccIndex = 0; ccIndex < rowContentControls.Count; ccIndex++)
            {
                var contentControl = rowContentControls[ccIndex];
                ExtractRowContentControlStyleAsync(contentControl, ccIndex, styles, patterns, templateId, createdBy);
            }
        }

        // COMPREHENSIVE DIRECT FORMATTING IMPLEMENTATION
        // Note: ExtractRunDirectFormatting method moved to DocumentProcessingServiceHelpers.cs for shared access

        // LIST STYLES IMPLEMENTATION
        private async Task ExtractListStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            // Extract numbering definitions
            var numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart?.Numbering != null)
            {
                var abstractNums = numberingPart.Numbering.Descendants<AbstractNum>().ToList();
                for (int numIndex = 0; numIndex < abstractNums.Count; numIndex++)
                {
                    var abstractNum = abstractNums[numIndex];
                    ExtractNumberingStyleAsync(abstractNum, numIndex, styles, templateId, createdBy);
                }
            }

            // Extract list paragraphs from document with enhanced direct formatting extraction
            var listParagraphs = mainPart.Document.Body?.Descendants<Paragraph>()
                .Where(p => p.ParagraphProperties?.NumberingProperties != null).ToList() ?? new List<Paragraph>();

            for (int listIndex = 0; listIndex < listParagraphs.Count; listIndex++)
            {
                var paragraph = listParagraphs[listIndex];
                // Now properly extracts direct formatting patterns from list item runs
                ExtractListParagraphStyleAsync(paragraph, listIndex, styles, patterns, templateId, createdBy);
            }
        }

        // IMAGE STYLES IMPLEMENTATION
        private void ExtractImageStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            // Extract drawings and images
            var drawings = mainPart.Document.Body?.Descendants<Drawing>().ToList() ?? new List<Drawing>();
            for (int drawingIndex = 0; drawingIndex < drawings.Count; drawingIndex++)
            {
                var drawing = drawings[drawingIndex];
                ExtractDrawingStyleAsync(drawing, drawingIndex, styles, templateId, createdBy);
            }

            // Extract inline pictures
            var pictures = mainPart.Document.Body?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().ToList() ?? new List<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
            for (int pictureIndex = 0; pictureIndex < pictures.Count; pictureIndex++)
            {
                var picture = pictures[pictureIndex];
                ExtractPictureStyleAsync(picture, pictureIndex, styles, templateId, createdBy);
            }
        }

        // SECTION STYLES IMPLEMENTATION
        private void ExtractSectionStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var sectionProperties = mainPart.Document.Body?.Descendants<SectionProperties>().ToList() ?? new List<SectionProperties>();
            
            for (int sectionIndex = 0; sectionIndex < sectionProperties.Count; sectionIndex++)
            {
                var sectionProps = sectionProperties[sectionIndex];
                ExtractSectionPropertiesStyleAsync(sectionProps, sectionIndex, styles, templateId, createdBy);
            }
        }

        // FIELD STYLES IMPLEMENTATION
        private void ExtractFieldStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var simpleFields = mainPart.Document.Body?.Descendants<SimpleField>().ToList() ?? new List<SimpleField>();
            for (int fieldIndex = 0; fieldIndex < simpleFields.Count; fieldIndex++)
            {
                var field = simpleFields[fieldIndex];
                ExtractSimpleFieldStyleAsync(field, fieldIndex, styles, templateId, createdBy);
            }

            var fieldChars = mainPart.Document.Body?.Descendants<FieldChar>().ToList() ?? new List<FieldChar>();
            for (int fieldIndex = 0; fieldIndex < fieldChars.Count; fieldIndex++)
            {
                var fieldChar = fieldChars[fieldIndex];
                ExtractComplexFieldStyleAsync(fieldChar, fieldIndex, styles, templateId, createdBy);
            }
        }

        // HYPERLINK STYLES IMPLEMENTATION
        private void ExtractHyperlinkStylesAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var hyperlinks = mainPart.Document.Body?.Descendants<Hyperlink>().ToList() ?? new List<Hyperlink>();
            for (int hyperlinkIndex = 0; hyperlinkIndex < hyperlinks.Count; hyperlinkIndex++)
            {
                var hyperlink = hyperlinks[hyperlinkIndex];
                ExtractHyperlinkStyleAsync(hyperlink, hyperlinkIndex, styles, patterns, templateId, createdBy);
            }
        }

        // STYLE DEFINITIONS IMPLEMENTATION
        private void ExtractStyleDefinitionsAsync(MainDocumentPart mainPart, List<TextStyle> styles, List<DirectFormatPattern> patterns, int templateId, string createdBy)
        {
            var stylesPart = mainPart.StyleDefinitionsPart;
            if (stylesPart?.Styles != null)
            {
                var documentStyles = stylesPart.Styles.Descendants<Style>().ToList();
                for (int styleIndex = 0; styleIndex < documentStyles.Count; styleIndex++)
                {
                    var style = documentStyles[styleIndex];
                    ExtractDocumentStyleAsync(style, styleIndex, styles, templateId, createdBy);
                }
            }
        }

        private async Task SaveExtractedStylesAsync(List<TextStyle> styles, List<DirectFormatPattern> patterns)
        {
            // Create a mapping from temporary IDs to actual TextStyle objects
            // We need to properly track which TextStyle has which temporary ID
            var tempIdToStyleMapping = new Dictionary<int, TextStyle>();
            
            // Build mapping by checking each style's temporary ID that was actually assigned
            foreach (var style in styles)
            {
                // During extraction, styles with temporary IDs have their Id property set temporarily
                // We need to find patterns that reference each style by examining the patterns
                var relatedPatterns = patterns.Where(p => p.TextStyleId < 0).ToList();
                foreach (var pattern in relatedPatterns)
                {
                    var tempId = pattern.TextStyleId;
                    if (tempId < 0 && !tempIdToStyleMapping.ContainsKey(tempId))
                    {
                        // Find the style this pattern should belong to by matching context
                        var matchingStyle = styles.FirstOrDefault(s => 
                            ShouldPatternBelongToStyle(pattern, s));
                        if (matchingStyle != null)
                        {
                            tempIdToStyleMapping[tempId] = matchingStyle;
                        }
                    }
                }
            }
            
            // Save TextStyles first to get their database IDs
            if (styles.Any())
            {
                _context.TextStyles.AddRange(styles);
                await _context.SaveChangesAsync(); // Save styles first to get IDs
            }
            
            // Now update DirectFormatPatterns with correct TextStyleId references
            if (patterns.Any())
            {
                var validPatterns = new List<DirectFormatPattern>();
                
                foreach (var pattern in patterns)
                {
                    // If pattern has a negative TextStyleId (temporary), map it to the real ID
                    if (pattern.TextStyleId < 0 && tempIdToStyleMapping.ContainsKey(pattern.TextStyleId))
                    {
                        var correspondingStyle = tempIdToStyleMapping[pattern.TextStyleId];
                        if (correspondingStyle.Id > 0)
                        {
                            pattern.TextStyleId = correspondingStyle.Id;
                            validPatterns.Add(pattern);
                        }
                        else
                        {
                            _logger.LogWarning($"TextStyle not properly saved for pattern: {pattern.PatternName}");
                        }
                    }
                    else if (pattern.TextStyleId > 0)
                    {
                        // Pattern already has a valid ID
                        validPatterns.Add(pattern);
                    }
                    else
                    {
                        _logger.LogWarning($"Invalid TextStyleId for DirectFormatPattern: {pattern.PatternName}");
                    }
                }
                
                if (validPatterns.Any())
                {
                    _context.DirectFormatPatterns.AddRange(validPatterns);
                    await _context.SaveChangesAsync();
                }
            }
        }

        /// <summary>
        /// Determines if a direct format pattern should belong to a specific style based on context
        /// </summary>
        public bool ShouldPatternBelongToStyle(DirectFormatPattern pattern, TextStyle style)
        {
            if (style.FormattingContext == null || string.IsNullOrEmpty(pattern.Context))
                return false;

            // Extract paragraph index from pattern context (e.g., "Paragraph:14,Run:0" -> 14)
            var contextParts = pattern.Context.Split(',');
            var paragraphPart = contextParts.FirstOrDefault(p => p.StartsWith("Paragraph:"));
            if (paragraphPart == null) return false;

            if (int.TryParse(paragraphPart.Split(':')[1], out var patternParagraphIndex))
            {
                // For paragraph styles, match by paragraph index
                if (style.StyleType == "Paragraph" && 
                    style.FormattingContext.ParagraphIndex == patternParagraphIndex)
                {
                    return true;
                }

                // For table cell paragraph styles, also check if pattern context matches
                if (style.StyleType == "TableCellParagraph")
                {
                    // Pattern should belong to table cell paragraph with matching indices
                    var contextKey = style.FormattingContext.ContextKey;
                    return contextKey.Contains($"Paragraph:{patternParagraphIndex}");
                }

                // For sections, check if this is a section-level pattern
                if (style.StyleType == "Section")
                {
                    // Section patterns typically come from paragraphs within that section
                    return style.FormattingContext.SectionIndex >= 0;
                }
            }

            return false;
        }

        // Interface implementations that delegate to extension methods
        public async Task<VerificationResultDto> VerifyDocumentAsync(Stream documentStream, string fileName, int templateId, string createdBy, bool strictMode = true, List<string>? ignoreStyleTypes = null)
        {
            return await DocumentProcessingServiceExtensions.VerifyDocumentAsync(this, documentStream, fileName, templateId, createdBy, strictMode, ignoreStyleTypes);
        }

        public async Task<List<TextStyle>> ExtractDocumentStylesAsync(Stream documentStream, string fileName)
        {
            return await DocumentProcessingServiceExtensions.ExtractDocumentStylesAsync(this, documentStream, fileName);
        }

        public List<MismatchReportDto> CompareStyles(List<TextStyle> templateStyles, List<TextStyle> documentStyles, bool strictMode = true)
        {
            return DocumentProcessingServiceExtensions.CompareStyles(this, templateStyles, documentStyles, strictMode);
        }

        public async Task<bool> ValidateDocumentFormatAsync(Stream documentStream)
        {
            return await DocumentProcessingServiceExtensions.ValidateDocumentFormatAsync(this, documentStream);
        }

        public async Task<DocumentMetadata> GetDocumentMetadataAsync(Stream documentStream)
        {
            return await DocumentProcessingServiceExtensions.GetDocumentMetadataAsync(this, documentStream);
        }
    }
} 