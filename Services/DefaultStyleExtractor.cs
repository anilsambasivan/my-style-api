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
            
            foreach (var style in styles)
            {
                var defaultStyle = CreateDefaultStyleFromStyle(style, templateId, createdBy);
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
        private static DefaultStyle? CreateDefaultStyleFromStyle(Style style, int templateId, string createdBy)
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

            // Extract font properties from run properties
            if (style.StyleRunProperties != null)
            {
                ExtractFontProperties(style.StyleRunProperties, defaultStyle);
            }

            // Extract paragraph properties
            if (style.StyleParagraphProperties != null)
            {
                ExtractParagraphProperties(style.StyleParagraphProperties, defaultStyle);
            }

            return defaultStyle;
        }

        /// <summary>
        /// Extracts font properties from StyleRunProperties
        /// </summary>
        private static void ExtractFontProperties(StyleRunProperties runProps, DefaultStyle defaultStyle)
        {
            if (runProps.RunFonts != null)
            {
                defaultStyle.FontFamily = runProps.RunFonts.Ascii?.Value ?? 
                                         runProps.RunFonts.HighAnsi?.Value ?? 
                                         runProps.RunFonts.ComplexScript?.Value ?? 
                                         runProps.RunFonts.EastAsia?.Value ?? 
                                         string.Empty;
            }

            if (runProps.FontSize != null && float.TryParse(runProps.FontSize.Val?.Value, out var fontSize))
            {
                defaultStyle.FontSize = fontSize / 2f; // Convert half-points to points
            }

            defaultStyle.IsBold = runProps.Bold != null && (runProps.Bold.Val?.Value ?? true);
            defaultStyle.IsItalic = runProps.Italic != null && (runProps.Italic.Val?.Value ?? true);
            defaultStyle.IsUnderline = runProps.Underline != null;

            if (runProps.Color != null)
            {
                defaultStyle.Color = runProps.Color.Val?.Value ?? string.Empty;
            }
        }

        /// <summary>
        /// Extracts paragraph properties from StyleParagraphProperties
        /// </summary>
        private static void ExtractParagraphProperties(StyleParagraphProperties paraProps, DefaultStyle defaultStyle)
        {
            if (paraProps.Justification != null)
            {
                defaultStyle.Alignment = paraProps.Justification.Val?.ToString() ?? string.Empty;
            }

            if (paraProps.SpacingBetweenLines != null)
            {
                if (paraProps.SpacingBetweenLines.Before != null && 
                    int.TryParse(paraProps.SpacingBetweenLines.Before.Value, out var spacingBefore))
                {
                    defaultStyle.SpacingBefore = ConvertTwipsToPoints(spacingBefore);
                }

                if (paraProps.SpacingBetweenLines.After != null && 
                    int.TryParse(paraProps.SpacingBetweenLines.After.Value, out var spacingAfter))
                {
                    defaultStyle.SpacingAfter = ConvertTwipsToPoints(spacingAfter);
                }

                if (paraProps.SpacingBetweenLines.Line != null && 
                    int.TryParse(paraProps.SpacingBetweenLines.Line.Value, out var lineSpacing))
                {
                    defaultStyle.LineSpacing = lineSpacing / 240f; // Convert to line spacing multiplier
                }
            }

            if (paraProps.Indentation != null)
            {
                if (paraProps.Indentation.Left != null && 
                    int.TryParse(paraProps.Indentation.Left.Value, out var leftIndent))
                {
                    defaultStyle.IndentationLeft = ConvertTwipsToPoints(leftIndent);
                }

                if (paraProps.Indentation.Right != null && 
                    int.TryParse(paraProps.Indentation.Right.Value, out var rightIndent))
                {
                    defaultStyle.IndentationRight = ConvertTwipsToPoints(rightIndent);
                }

                if (paraProps.Indentation.FirstLine != null && 
                    int.TryParse(paraProps.Indentation.FirstLine.Value, out var firstLineIndent))
                {
                    defaultStyle.FirstLineIndent = ConvertTwipsToPoints(firstLineIndent);
                }
            }
        }

        /// <summary>
        /// Converts twips to points (1 point = 20 twips)
        /// </summary>
        private static float ConvertTwipsToPoints(int twips)
        {
            return twips / 20f;
        }
    }
} 