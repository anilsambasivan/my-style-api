using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocStyleVerify.API.Models;
using System.Xml;

namespace DocStyleVerify.API.Services
{
    /// <summary>
    /// Service for extracting default styles from styles.xml
    /// </summary>
    public static class DefaultStyleExtractor
    {
        /// <summary>
        /// Extracts default styles from the document's styles.xml
        /// </summary>
        /// <param name="mainPart">The main document part</param>
        /// <param name="templateId">Template ID</param>
        /// <param name="createdBy">User who initiated the extraction</param>
        /// <returns>List of default styles</returns>
        public static List<DefaultStyle> ExtractDefaultStyles(MainDocumentPart mainPart, int templateId, string createdBy)
        {
            var defaultStyles = new List<DefaultStyle>();
            
            var stylesPart = mainPart.StyleDefinitionsPart;
            if (stylesPart?.Styles == null)
                return defaultStyles;

            var styles = stylesPart.Styles.Descendants<Style>().ToList();
            
            // Create a dictionary for quick style lookups for inheritance
            var styleDict = styles.ToDictionary(s => s.StyleId?.Value ?? "", s => s);
            
            foreach (var style in styles)
            {
                var defaultStyle = CreateDefaultStyleFromStyle(style, templateId, createdBy, styleDict);
                if (defaultStyle != null)
                {
                    defaultStyles.Add(defaultStyle);
                }
            }

            return defaultStyles;
        }

        /// <summary>
        /// Creates a DefaultStyle entity from a Style element
        /// </summary>
        private static DefaultStyle? CreateDefaultStyleFromStyle(Style style, int templateId, string createdBy, Dictionary<string, Style> styleDict)
        {
            if (style.StyleId?.Value == null)
                return null;

            var defaultStyle = new DefaultStyle
            {
                TemplateId = templateId,
                StyleId = style.StyleId.Value,
                Name = style.StyleName?.Val?.Value ?? style.StyleId.Value,
                Type = style.Type?.ToString() ?? "Unknown",
                BasedOn = style.BasedOn?.Val?.Value ?? string.Empty,
                NextStyle = style.NextParagraphStyle?.Val?.Value ?? string.Empty,
                IsDefault = style.Default?.Value ?? false,
                IsCustom = style.CustomStyle?.Value ?? false,
                Priority = style.UIPriority?.Val?.Value ?? 0,
                IsHidden = style.SemiHidden != null,
                IsQuickStyle = style.PrimaryStyle != null,
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                RawXml = style.OuterXml
            };

            // Extract font properties from run properties with inheritance
            ExtractFontPropertiesWithInheritance(style, defaultStyle, styleDict);

            // Extract paragraph properties with inheritance
            ExtractParagraphPropertiesWithInheritance(style, defaultStyle, styleDict);

            return defaultStyle;
        }

        /// <summary>
        /// Extracts font properties with inheritance from base styles
        /// </summary>
        private static void ExtractFontPropertiesWithInheritance(Style style, DefaultStyle defaultStyle, Dictionary<string, Style> styleDict)
        {
            // Start with the current style and walk up the inheritance chain
            var currentStyle = style;
            var visitedStyles = new HashSet<string>(); // Prevent infinite loops
            
            while (currentStyle != null)
            {
                // Prevent infinite loops in case of circular references
                if (visitedStyles.Contains(currentStyle.StyleId?.Value ?? ""))
                    break;
                    
                visitedStyles.Add(currentStyle.StyleId?.Value ?? "");
                
                if (currentStyle.StyleRunProperties != null)
                {
                    // Extract font family (including theme fonts)
                    if (string.IsNullOrEmpty(defaultStyle.FontFamily))
                    {
                        defaultStyle.FontFamily = ExtractFontFamily(currentStyle.StyleRunProperties);
                    }
                    
                    // Extract font size
                    if (defaultStyle.FontSize == 0)
                    {
                        var fontSize = ExtractFontSize(currentStyle.StyleRunProperties);
                        if (fontSize > 0)
                            defaultStyle.FontSize = fontSize;
                    }
                    
                    // Extract boolean properties (only if not already set)
                    if (!defaultStyle.IsBold && currentStyle.StyleRunProperties.Bold != null)
                        defaultStyle.IsBold = currentStyle.StyleRunProperties.Bold.Val?.Value ?? true;
                        
                    if (!defaultStyle.IsItalic && currentStyle.StyleRunProperties.Italic != null)
                        defaultStyle.IsItalic = currentStyle.StyleRunProperties.Italic.Val?.Value ?? true;
                        
                    if (!defaultStyle.IsUnderline && currentStyle.StyleRunProperties.Underline != null)
                        defaultStyle.IsUnderline = true;
                    
                    // Extract color
                    if (string.IsNullOrEmpty(defaultStyle.Color))
                    {
                        var color = ExtractColor(currentStyle.StyleRunProperties);
                        if (!string.IsNullOrEmpty(color))
                            defaultStyle.Color = color;
                    }
                }
                
                // Move to base style
                var basedOnStyleId = currentStyle.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOnStyleId) || !styleDict.ContainsKey(basedOnStyleId))
                    break;
                    
                currentStyle = styleDict[basedOnStyleId];
            }
            
            // Apply defaults if still not set
            if (string.IsNullOrEmpty(defaultStyle.FontFamily))
                defaultStyle.FontFamily = "Calibri"; // Word default
                
            if (defaultStyle.FontSize == 0)
                defaultStyle.FontSize = 11; // Word default (22 half-points)
                
            if (string.IsNullOrEmpty(defaultStyle.Color))
                defaultStyle.Color = "000000"; // Black
        }

        /// <summary>
        /// Extracts paragraph properties with inheritance from base styles
        /// </summary>
        private static void ExtractParagraphPropertiesWithInheritance(Style style, DefaultStyle defaultStyle, Dictionary<string, Style> styleDict)
        {
            var currentStyle = style;
            var visitedStyles = new HashSet<string>();
            
            while (currentStyle != null)
            {
                if (visitedStyles.Contains(currentStyle.StyleId?.Value ?? ""))
                    break;
                    
                visitedStyles.Add(currentStyle.StyleId?.Value ?? "");
                
                if (currentStyle.StyleParagraphProperties != null)
                {
                    var paraProps = currentStyle.StyleParagraphProperties;
                    
                    // Extract alignment
                    if (string.IsNullOrEmpty(defaultStyle.Alignment) && paraProps.Justification != null)
                    {
                        defaultStyle.Alignment = paraProps.Justification.Val?.ToString() ?? string.Empty;
                    }
                    
                    // Extract spacing
                    if (paraProps.SpacingBetweenLines != null)
                    {
                        if (defaultStyle.SpacingBefore == 0 && paraProps.SpacingBetweenLines.Before != null && 
                            int.TryParse(paraProps.SpacingBetweenLines.Before.Value, out var spacingBefore))
                        {
                            defaultStyle.SpacingBefore = ConvertTwipsToPoints(spacingBefore);
                        }

                        if (defaultStyle.SpacingAfter == 0 && paraProps.SpacingBetweenLines.After != null && 
                            int.TryParse(paraProps.SpacingBetweenLines.After.Value, out var spacingAfter))
                        {
                            defaultStyle.SpacingAfter = ConvertTwipsToPoints(spacingAfter);
                        }

                        if (defaultStyle.LineSpacing == 0 && paraProps.SpacingBetweenLines.Line != null && 
                            int.TryParse(paraProps.SpacingBetweenLines.Line.Value, out var lineSpacing))
                        {
                            defaultStyle.LineSpacing = lineSpacing / 240f;
                        }
                    }
                    
                    // Extract indentation
                    if (paraProps.Indentation != null)
                    {
                        if (defaultStyle.IndentationLeft == 0 && paraProps.Indentation.Left != null && 
                            int.TryParse(paraProps.Indentation.Left.Value, out var leftIndent))
                        {
                            defaultStyle.IndentationLeft = ConvertTwipsToPoints(leftIndent);
                        }

                        if (defaultStyle.IndentationRight == 0 && paraProps.Indentation.Right != null && 
                            int.TryParse(paraProps.Indentation.Right.Value, out var rightIndent))
                        {
                            defaultStyle.IndentationRight = ConvertTwipsToPoints(rightIndent);
                        }

                        if (defaultStyle.FirstLineIndent == 0 && paraProps.Indentation.FirstLine != null && 
                            int.TryParse(paraProps.Indentation.FirstLine.Value, out var firstLineIndent))
                        {
                            defaultStyle.FirstLineIndent = ConvertTwipsToPoints(firstLineIndent);
                        }
                    }
                }
                
                // Move to base style
                var basedOnStyleId = currentStyle.BasedOn?.Val?.Value;
                if (string.IsNullOrEmpty(basedOnStyleId) || !styleDict.ContainsKey(basedOnStyleId))
                    break;
                    
                currentStyle = styleDict[basedOnStyleId];
            }
        }

        /// <summary>
        /// Extracts font family including theme fonts
        /// </summary>
        private static string ExtractFontFamily(StyleRunProperties runProps)
        {
            if (runProps.RunFonts == null)
                return string.Empty;
                
            // First try direct font values
            var fontName = runProps.RunFonts.Ascii?.Value ?? 
                          runProps.RunFonts.HighAnsi?.Value ?? 
                          runProps.RunFonts.ComplexScript?.Value ?? 
                          runProps.RunFonts.EastAsia?.Value;
                          
            if (!string.IsNullOrEmpty(fontName))
                return ResolveThemeFont(fontName);
            
            // Then try theme fonts
            var themeFontName = runProps.RunFonts.AsciiTheme?.ToString() ?? 
                               runProps.RunFonts.HighAnsiTheme?.ToString() ?? 
                               runProps.RunFonts.EastAsiaTheme?.ToString() ?? 
                               runProps.RunFonts.ComplexScriptTheme?.ToString();
                               
            if (!string.IsNullOrEmpty(themeFontName))
                return ResolveThemeFont(themeFontName);
                
            return string.Empty;
        }

        /// <summary>
        /// Extracts font size from run properties
        /// </summary>
        private static float ExtractFontSize(StyleRunProperties runProps)
        {
            if (runProps.FontSize != null && float.TryParse(runProps.FontSize.Val?.Value, out var fontSize))
            {
                return fontSize / 2f; // Convert half-points to points
            }
            return 0;
        }

        /// <summary>
        /// Extracts color from run properties
        /// </summary>
        private static string ExtractColor(StyleRunProperties runProps)
        {
            if (runProps.Color != null)
            {
                var colorValue = runProps.Color.Val?.Value;
                if (!string.IsNullOrEmpty(colorValue) && colorValue != "auto")
                {
                    return ConvertColorToHex(colorValue);
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// Resolves theme fonts to actual font names
        /// </summary>
        private static string ResolveThemeFont(string themeFontValue)
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
                { "minorfont", "Calibri" },   // Minor theme font
                { "majorfont", "Cambria" },   // Major theme font
                { "minorhansi", "Calibri" },  // Minor H-ANSI theme font
                { "majorhansi", "Cambria" },  // Major H-ANSI theme font
                { "minoreastasia", "Calibri" }, // Minor East Asian theme font
                { "majoreastasia", "Cambria" }, // Major East Asian theme font
                { "minorbidi", "Calibri" },   // Minor Bidi theme font
                { "majorbidi", "Cambria" }    // Major Bidi theme font
            };

            return themeFonts.ContainsKey(themeFontValue.ToLower()) ? 
                   themeFonts[themeFontValue.ToLower()] : 
                   themeFontValue; // Return original if not a theme font
        }

        /// <summary>
        /// Converts color value to hex format
        /// </summary>
        private static string ConvertColorToHex(string colorValue)
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

        /// <summary>
        /// Resolves theme colors to hex values
        /// </summary>
        private static string ResolveThemeColor(string themeColorValue)
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

        /// <summary>
        /// Converts twips to points
        /// </summary>
        private static float ConvertTwipsToPoints(int twips)
        {
            return twips / 20f; // 20 twips = 1 point
        }
    }
} 