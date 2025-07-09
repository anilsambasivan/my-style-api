using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocStyleVerify.API.Models;
using DocStyleVerify.API.DTOs;
using DocStyleVerify.API.Data;
using Microsoft.EntityFrameworkCore;
using System.Text.Json;

namespace DocStyleVerify.API.Services
{
    /// <summary>
    /// Enhanced verification extensions for DocumentProcessingService to handle direct formatting patterns
    /// </summary>
    public static class DocumentProcessingServiceEnhanced
    {
        /// <summary>
        /// Enhanced verification that properly handles direct formatting patterns and style combinations
        /// </summary>
        public static async Task<VerificationResultDto> VerifyDocumentWithEnhancedLogicAsync(
            this DocumentProcessingService service,
            Stream documentStream,
            string fileName,
            int templateId,
            string createdBy,
            bool strictMode = true,
            List<string>? ignoreStyleTypes = null)
        {
            var result = new VerificationResultDto
            {
                TemplateId = templateId,
                DocumentName = fileName,
                DocumentPath = fileName,
                VerificationDate = DateTime.UtcNow,
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                Status = "Pending"
            };

            try
            {
                // Get template styles from database with direct format patterns
                var templateStyles = await GetTemplateStylesWithPatternsAsync(service, templateId);
                if (!templateStyles.Any())
                {
                    result.Status = "Failed";
                    result.ErrorMessage = "No template styles found for comparison";
                    return result;
                }

                // Extract document styles with patterns
                var (documentStyles, documentPatterns) = await service.ExtractStylesOnlyAsync(documentStream, 0, createdBy);

                // Enhanced comparison that handles direct formatting properly
                var mismatches = CompareStylesEnhanced(templateStyles, documentStyles, strictMode, ignoreStyleTypes);

                // Map mismatches to DTOs
                result.Mismatches = mismatches.Select(m => new MismatchDto
                {
                    ContextKey = m.ContextKey,
                    Location = m.Location,
                    StructuralRole = m.StructuralRole,
                    Expected = m.Expected,
                    Actual = m.Actual,
                    MismatchFields = m.MismatchedFields,
                    SampleText = m.SampleText,
                    Severity = m.Severity,
                    CreatedOn = DateTime.UtcNow
                }).ToList();

                result.TotalMismatches = result.Mismatches.Count;
                result.Status = "Completed";

                // Save to database
                await SaveVerificationResultAsync(service, result, templateId);
            }
            catch (Exception ex)
            {
                result.Status = "Failed";
                result.ErrorMessage = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// Enhanced style comparison that properly handles direct formatting combinations
        /// </summary>
        public static List<MismatchReportDto> CompareStylesEnhanced(
            List<TextStyle> templateStyles,
            List<TextStyle> documentStyles,
            bool strictMode = true,
            List<string>? ignoreStyleTypes = null)
        {
            var mismatches = new List<MismatchReportDto>();
            ignoreStyleTypes ??= new List<string>();

            // Group styles by their effective location/context
            var templateStylesByContext = GroupStylesByContext(templateStyles);
            var documentStylesByContext = GroupStylesByContext(documentStyles);

            foreach (var templateGroup in templateStylesByContext)
            {
                var contextKey = templateGroup.Key;
                var templateStyle = templateGroup.Value;

                // Skip ignored style types
                if (ignoreStyleTypes.Contains(templateStyle.StyleType))
                    continue;

                var documentStyle = documentStylesByContext.GetValueOrDefault(contextKey);

                if (documentStyle == null)
                {
                    // Style missing in document
                    mismatches.Add(CreateMissingStyleMismatch(templateStyle, contextKey));
                }
                else
                {
                    // Compare styles with enhanced logic for direct formatting
                    var styleMismatches = CompareStylesWithDirectFormatting(templateStyle, documentStyle, contextKey, strictMode);
                    mismatches.AddRange(styleMismatches);
                }
            }

            // Check for extra styles in document that aren't in template
            foreach (var documentGroup in documentStylesByContext)
            {
                var contextKey = documentGroup.Key;
                var documentStyle = documentGroup.Value;

                if (!templateStylesByContext.ContainsKey(contextKey) && 
                    !ignoreStyleTypes.Contains(documentStyle.StyleType))
                {
                    // Check if this is a structural element that should exist but has wrong formatting
                    var isStructuralElement = IsStructuralElementEnhanced(documentStyle);
                    
                    if (isStructuralElement)
                    {
                        // This is a table cell, paragraph, etc. that exists but has unexpected formatting
                        // Report this as extra formatting rather than "should not exist"
                        mismatches.Add(CreateUnexpectedFormattingMismatchEnhanced(documentStyle, contextKey));
                    }
                    else
                    {
                        // This is truly an extra style that shouldn't be there
                        mismatches.Add(CreateExtraStyleMismatch(documentStyle, contextKey));
                    }
                }
            }

            // Filter out mismatches where SampleText is empty or null
            // This reduces noise by focusing only on content that actually exists
            return mismatches.Where(m => !string.IsNullOrWhiteSpace(m.SampleText)).ToList();
        }

        /// <summary>
        /// Compares individual styles considering base style + direct formatting combinations
        /// </summary>
        private static List<MismatchReportDto> CompareStylesWithDirectFormatting(
            TextStyle templateStyle,
            TextStyle documentStyle,
            string contextKey,
            bool strictMode)
        {
            var mismatches = new List<MismatchReportDto>();

            // Get the effective style properties (base style + direct formatting)
            var templateEffective = GetEffectiveStyleProperties(templateStyle);
            var documentEffective = GetEffectiveStyleProperties(documentStyle);

            // Find differences in effective properties
            var differences = FindStyleDifferences(templateEffective, documentEffective);

            if (differences.Any())
            {
                // Group differences by severity and create detailed mismatches
                var mismatch = new MismatchReportDto
                {
                    ContextKey = contextKey,
                    Location = GenerateEnhancedLocation(templateStyle.FormattingContext, documentStyle.FormattingContext),
                    StructuralRole = templateStyle.FormattingContext?.StructuralRole ?? documentStyle.FormattingContext?.StructuralRole ?? "Unknown",
                    Expected = templateEffective,
                    Actual = documentEffective,
                    MismatchedFields = differences,
                    SampleText = documentStyle.FormattingContext?.SampleText ?? templateStyle.FormattingContext?.SampleText ?? "",
                    Severity = DetermineEnhancedSeverity(differences, templateStyle, documentStyle),
                    RecommendedAction = GenerateEnhancedRecommendedAction(differences, templateStyle, documentStyle)
                };

                mismatches.Add(mismatch);

                // Add specific mismatches for each direct formatting deviation
                var directFormattingMismatches = CompareDirectFormattingPatterns(templateStyle, documentStyle, contextKey);
                mismatches.AddRange(directFormattingMismatches);
            }

            return mismatches;
        }

        /// <summary>
        /// Compares direct formatting patterns between template and document
        /// </summary>
        private static List<MismatchReportDto> CompareDirectFormattingPatterns(
            TextStyle templateStyle,
            TextStyle documentStyle,
            string contextKey)
        {
            var mismatches = new List<MismatchReportDto>();

            // Group template patterns by context
            var templatePatterns = templateStyle.DirectFormatPatterns.ToLookup(p => p.Context);
            var documentPatterns = documentStyle.DirectFormatPatterns.ToLookup(p => p.Context);

            // Check each template pattern
            foreach (var templatePatternGroup in templatePatterns)
            {
                var patternContext = templatePatternGroup.Key;
                var templatePattern = templatePatternGroup.First(); // Take first if multiple

                var documentPattern = documentPatterns[patternContext].FirstOrDefault();

                if (documentPattern == null)
                {
                    // Direct formatting pattern missing in document
                    mismatches.Add(CreateMissingDirectFormattingMismatch(templatePattern, contextKey, patternContext));
                }
                else
                {
                    // Compare direct formatting properties
                    var patternDifferences = CompareDirectFormattingProperties(templatePattern, documentPattern);
                    if (patternDifferences.Any())
                    {
                        mismatches.Add(CreateDirectFormattingMismatch(templatePattern, documentPattern, contextKey, patternContext, patternDifferences));
                    }
                }
            }

            // Check for extra direct formatting in document
            foreach (var documentPatternGroup in documentPatterns)
            {
                var patternContext = documentPatternGroup.Key;
                
                if (!templatePatterns.Contains(patternContext))
                {
                    var documentPattern = documentPatternGroup.First();
                    mismatches.Add(CreateExtraDirectFormattingMismatch(documentPattern, contextKey, patternContext));
                }
            }

            return mismatches;
        }

        /// <summary>
        /// Gets effective style properties combining base style and direct formatting
        /// </summary>
        private static Dictionary<string, object> GetEffectiveStyleProperties(TextStyle style)
        {
            var properties = new Dictionary<string, object>();

            // Base style properties
            AddPropertyIfNotDefault(properties, "FontFamily", style.FontFamily);
            AddPropertyIfNotDefault(properties, "FontSize", style.FontSize);
            AddPropertyIfNotDefault(properties, "IsBold", style.IsBold);
            AddPropertyIfNotDefault(properties, "IsItalic", style.IsItalic);
            AddPropertyIfNotDefault(properties, "IsUnderline", style.IsUnderline);
            AddPropertyIfNotDefault(properties, "IsStrikethrough", style.IsStrikethrough);
            AddPropertyIfNotDefault(properties, "Color", style.Color);
            AddPropertyIfNotDefault(properties, "Alignment", style.Alignment);
            AddPropertyIfNotDefault(properties, "SpacingBefore", style.SpacingBefore);
            AddPropertyIfNotDefault(properties, "SpacingAfter", style.SpacingAfter);
            AddPropertyIfNotDefault(properties, "IndentationLeft", style.IndentationLeft);
            AddPropertyIfNotDefault(properties, "IndentationRight", style.IndentationRight);
            AddPropertyIfNotDefault(properties, "LineSpacing", style.LineSpacing);

            // Apply direct formatting overrides
            foreach (var pattern in style.DirectFormatPatterns)
            {
                ApplyDirectFormattingToProperties(properties, pattern);
            }

            return properties;
        }

        /// <summary>
        /// Applies direct formatting pattern to style properties
        /// </summary>
        private static void ApplyDirectFormattingToProperties(Dictionary<string, object> properties, DirectFormatPattern pattern)
        {
            if (!string.IsNullOrEmpty(pattern.FontFamily))
                properties["FontFamily"] = pattern.FontFamily;
            
            if (pattern.FontSize.HasValue)
                properties["FontSize"] = pattern.FontSize.Value;
            
            if (pattern.IsBold.HasValue)
                properties["IsBold"] = pattern.IsBold.Value;
            
            if (pattern.IsItalic.HasValue)
                properties["IsItalic"] = pattern.IsItalic.Value;
            
            if (pattern.IsUnderline.HasValue)
                properties["IsUnderline"] = pattern.IsUnderline.Value;
            
            if (pattern.IsStrikethrough.HasValue)
                properties["IsStrikethrough"] = pattern.IsStrikethrough.Value;
            
            if (!string.IsNullOrEmpty(pattern.Color))
                properties["Color"] = pattern.Color;
            
            if (!string.IsNullOrEmpty(pattern.Alignment))
                properties["Alignment"] = pattern.Alignment;
            
            if (pattern.SpacingBefore.HasValue)
                properties["SpacingBefore"] = pattern.SpacingBefore.Value;
            
            if (pattern.SpacingAfter.HasValue)
                properties["SpacingAfter"] = pattern.SpacingAfter.Value;
            
            if (pattern.IndentationLeft.HasValue)
                properties["IndentationLeft"] = pattern.IndentationLeft.Value;
            
            if (pattern.IndentationRight.HasValue)
                properties["IndentationRight"] = pattern.IndentationRight.Value;
            
            if (pattern.LineSpacing.HasValue)
                properties["LineSpacing"] = pattern.LineSpacing.Value;
        }

        /// <summary>
        /// Finds differences between two sets of style properties
        /// </summary>
        private static List<string> FindStyleDifferences(Dictionary<string, object> template, Dictionary<string, object> document)
        {
            var differences = new List<string>();

            // Check all template properties
            foreach (var templateProp in template)
            {
                if (!document.ContainsKey(templateProp.Key))
                {
                    differences.Add(templateProp.Key);
                }
                else if (!AreValuesEqual(templateProp.Value, document[templateProp.Key]))
                {
                    differences.Add(templateProp.Key);
                }
            }

            // Check for extra properties in document
            foreach (var documentProp in document)
            {
                if (!template.ContainsKey(documentProp.Key))
                {
                    differences.Add(documentProp.Key);
                }
            }

            return differences;
        }

        /// <summary>
        /// Compares direct formatting pattern properties
        /// </summary>
        private static List<string> CompareDirectFormattingProperties(DirectFormatPattern template, DirectFormatPattern document)
        {
            var differences = new List<string>();

            if (template.FontFamily != document.FontFamily) differences.Add("FontFamily");
            if (template.FontSize != document.FontSize) differences.Add("FontSize");
            if (template.IsBold != document.IsBold) differences.Add("IsBold");
            if (template.IsItalic != document.IsItalic) differences.Add("IsItalic");
            if (template.IsUnderline != document.IsUnderline) differences.Add("IsUnderline");
            if (template.IsStrikethrough != document.IsStrikethrough) differences.Add("IsStrikethrough");
            if (template.Color != document.Color) differences.Add("Color");
            if (template.Alignment != document.Alignment) differences.Add("Alignment");
            if (template.SpacingBefore != document.SpacingBefore) differences.Add("SpacingBefore");
            if (template.SpacingAfter != document.SpacingAfter) differences.Add("SpacingAfter");
            if (template.IndentationLeft != document.IndentationLeft) differences.Add("IndentationLeft");
            if (template.IndentationRight != document.IndentationRight) differences.Add("IndentationRight");
            if (template.LineSpacing != document.LineSpacing) differences.Add("LineSpacing");

            return differences;
        }

        /// <summary>
        /// Groups styles by their effective context/location
        /// </summary>
        private static Dictionary<string, TextStyle> GroupStylesByContext(List<TextStyle> styles)
        {
            var grouped = new Dictionary<string, TextStyle>();

            foreach (var style in styles)
            {
                var contextKey = style.FormattingContext?.ContextKey ?? $"{style.StyleType}:{style.Name}";
                
                // If multiple styles exist for same context, prefer the one with direct formatting
                if (!grouped.ContainsKey(contextKey) || style.DirectFormatPatterns.Any())
                {
                    grouped[contextKey] = style;
                }
            }

            return grouped;
        }

        /// <summary>
        /// Creates a mismatch for missing style
        /// </summary>
        private static MismatchReportDto CreateMissingStyleMismatch(TextStyle templateStyle, string contextKey)
        {
            return new MismatchReportDto
            {
                ContextKey = contextKey,
                Location = GenerateEnhancedLocation(templateStyle.FormattingContext, null),
                StructuralRole = templateStyle.FormattingContext?.StructuralRole ?? "Unknown",
                Expected = GetEffectiveStyleProperties(templateStyle),
                Actual = new Dictionary<string, object> { { "Status", "Missing" } },
                MismatchedFields = new List<string> { "EntireStyle" },
                SampleText = templateStyle.FormattingContext?.SampleText ?? "",
                Severity = "High",
                RecommendedAction = $"Apply the expected style '{templateStyle.Name}' to this location"
            };
        }

        /// <summary>
        /// Creates a mismatch for extra style not in template
        /// </summary>
        private static MismatchReportDto CreateExtraStyleMismatch(TextStyle documentStyle, string contextKey)
        {
            return new MismatchReportDto
            {
                ContextKey = contextKey,
                Location = GenerateEnhancedLocation(null, documentStyle.FormattingContext),
                StructuralRole = documentStyle.FormattingContext?.StructuralRole ?? "Unknown",
                Expected = new Dictionary<string, object> { { "Status", "Should not exist" } },
                Actual = GetEffectiveStyleProperties(documentStyle),
                MismatchedFields = new List<string> { "EntireStyle" },
                SampleText = documentStyle.FormattingContext?.SampleText ?? "",
                Severity = "Medium",
                RecommendedAction = $"Remove the unexpected style '{documentStyle.Name}' from this location"
            };
        }

        /// <summary>
        /// Creates a mismatch for missing direct formatting
        /// </summary>
        private static MismatchReportDto CreateMissingDirectFormattingMismatch(DirectFormatPattern templatePattern, string contextKey, string patternContext)
        {
            return new MismatchReportDto
            {
                ContextKey = $"{contextKey}:{patternContext}",
                Location = $"{contextKey} > {patternContext}",
                StructuralRole = "DirectFormatting",
                Expected = SerializeDirectFormattingPattern(templatePattern),
                Actual = new Dictionary<string, object> { { "Status", "Missing" } },
                MismatchedFields = new List<string> { "DirectFormatting" },
                SampleText = templatePattern.SampleText,
                Severity = "Medium",
                RecommendedAction = $"Apply the missing direct formatting pattern '{templatePattern.PatternName}'"
            };
        }

        /// <summary>
        /// Creates a mismatch for direct formatting differences
        /// </summary>
        private static MismatchReportDto CreateDirectFormattingMismatch(
            DirectFormatPattern templatePattern, 
            DirectFormatPattern documentPattern, 
            string contextKey, 
            string patternContext, 
            List<string> differences)
        {
            return new MismatchReportDto
            {
                ContextKey = $"{contextKey}:{patternContext}",
                Location = $"{contextKey} > {patternContext}",
                StructuralRole = "DirectFormatting",
                Expected = SerializeDirectFormattingPattern(templatePattern),
                Actual = SerializeDirectFormattingPattern(documentPattern),
                MismatchedFields = differences,
                SampleText = documentPattern.SampleText,
                Severity = DetermineDirectFormattingSeverity(differences),
                RecommendedAction = $"Correct the direct formatting pattern '{documentPattern.PatternName}' to match template"
            };
        }

        /// <summary>
        /// Creates a mismatch for extra direct formatting
        /// </summary>
        private static MismatchReportDto CreateExtraDirectFormattingMismatch(DirectFormatPattern documentPattern, string contextKey, string patternContext)
        {
            return new MismatchReportDto
            {
                ContextKey = $"{contextKey}:{patternContext}",
                Location = $"{contextKey} > {patternContext}",
                StructuralRole = "DirectFormatting",
                Expected = new Dictionary<string, object> { { "Status", "Should not exist" } },
                Actual = SerializeDirectFormattingPattern(documentPattern),
                MismatchedFields = new List<string> { "DirectFormatting" },
                SampleText = documentPattern.SampleText,
                Severity = "High",
                RecommendedAction = $"Remove the unexpected direct formatting pattern '{documentPattern.PatternName}'"
            };
        }

        /// <summary>
        /// Helper methods
        /// </summary>
        private static void AddPropertyIfNotDefault(Dictionary<string, object> properties, string key, object? value)
        {
            if (value != null && !IsDefaultValue(value))
            {
                properties[key] = value;
            }
        }

        private static bool IsDefaultValue(object value)
        {
            return value switch
            {
                string s => string.IsNullOrEmpty(s),
                float f => f == 0,
                int i => i == 0,
                bool b => !b,
                _ => false
            };
        }

        private static bool AreValuesEqual(object? value1, object? value2)
        {
            if (value1 == null && value2 == null) return true;
            if (value1 == null || value2 == null) return false;
            
            return value1.Equals(value2);
        }

        private static string GenerateEnhancedLocation(FormattingContext? templateContext, FormattingContext? documentContext)
        {
            var context = templateContext ?? documentContext;
            if (context == null) return "Unknown";

            var parts = new List<string>();
            
            if (context.SectionIndex >= 0) parts.Add($"Section {context.SectionIndex + 1}");
            if (context.TableIndex >= 0) parts.Add($"Table {context.TableIndex + 1}");
            if (context.RowIndex >= 0) parts.Add($"Row {context.RowIndex + 1}");
            if (context.CellIndex >= 0) parts.Add($"Cell {context.CellIndex + 1}");
            if (context.ParagraphIndex >= 0) parts.Add($"Paragraph {context.ParagraphIndex + 1}");
            if (context.RunIndex >= 0) parts.Add($"Run {context.RunIndex + 1}");

            return string.Join(", ", parts);
        }

        private static string DetermineEnhancedSeverity(List<string> differences, TextStyle templateStyle, TextStyle documentStyle)
        {
            if (differences.Contains("FontFamily") || differences.Contains("Color"))
                return "Critical";
            if (differences.Contains("FontSize") || differences.Contains("IsBold") || differences.Contains("IsItalic"))
                return "High";
            if (differences.Contains("Alignment") || differences.Contains("LineSpacing"))
                return "Medium";
            return "Low";
        }

        private static string DetermineDirectFormattingSeverity(List<string> differences)
        {
            if (differences.Contains("FontFamily") || differences.Contains("Color"))
                return "High";
            if (differences.Contains("FontSize") || differences.Contains("IsBold") || differences.Contains("IsItalic"))
                return "Medium";
            return "Low";
        }

        private static string GenerateEnhancedRecommendedAction(List<string> differences, TextStyle templateStyle, TextStyle documentStyle)
        {
            if (differences.Count == 1)
            {
                var field = differences[0];
                return $"Change {field} from current value to match template style '{templateStyle.Name}'";
            }
            
            return $"Correct the following properties to match template style '{templateStyle.Name}': {string.Join(", ", differences)}";
        }

        private static Dictionary<string, object> SerializeDirectFormattingPattern(DirectFormatPattern pattern)
        {
            var properties = new Dictionary<string, object>();
            
            if (!string.IsNullOrEmpty(pattern.FontFamily)) properties["FontFamily"] = pattern.FontFamily;
            if (pattern.FontSize.HasValue) properties["FontSize"] = pattern.FontSize.Value;
            if (pattern.IsBold.HasValue) properties["IsBold"] = pattern.IsBold.Value;
            if (pattern.IsItalic.HasValue) properties["IsItalic"] = pattern.IsItalic.Value;
            if (pattern.IsUnderline.HasValue) properties["IsUnderline"] = pattern.IsUnderline.Value;
            if (pattern.IsStrikethrough.HasValue) properties["IsStrikethrough"] = pattern.IsStrikethrough.Value;
            if (!string.IsNullOrEmpty(pattern.Color)) properties["Color"] = pattern.Color;
            if (!string.IsNullOrEmpty(pattern.Alignment)) properties["Alignment"] = pattern.Alignment;
            if (pattern.SpacingBefore.HasValue) properties["SpacingBefore"] = pattern.SpacingBefore.Value;
            if (pattern.SpacingAfter.HasValue) properties["SpacingAfter"] = pattern.SpacingAfter.Value;
            if (pattern.IndentationLeft.HasValue) properties["IndentationLeft"] = pattern.IndentationLeft.Value;
            if (pattern.IndentationRight.HasValue) properties["IndentationRight"] = pattern.IndentationRight.Value;
            if (pattern.LineSpacing.HasValue) properties["LineSpacing"] = pattern.LineSpacing.Value;
            
            return properties;
        }

        /// <summary>
        /// Gets template styles with their direct formatting patterns from database
        /// </summary>
        private static async Task<List<TextStyle>> GetTemplateStylesWithPatternsAsync(
            DocumentProcessingService service,
            int templateId)
        {
            var contextProperty = typeof(DocumentProcessingService)
                .GetField("_context", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            if (contextProperty?.GetValue(service) is DocStyleVerifyDbContext context)
            {
                return await context.TextStyles
                    .Include(ts => ts.FormattingContext)
                    .Include(ts => ts.DirectFormatPatterns)
                    .Include(ts => ts.TabStops)
                    .Where(ts => ts.TemplateId == templateId)
                    .ToListAsync();
            }

            return new List<TextStyle>();
        }

        /// <summary>
        /// Saves verification result to database
        /// </summary>
        private static async Task SaveVerificationResultAsync(
            DocumentProcessingService service,
            VerificationResultDto result,
            int templateId)
        {
            var contextProperty = typeof(DocumentProcessingService)
                .GetField("_context", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

            if (contextProperty?.GetValue(service) is DocStyleVerifyDbContext context)
            {
                var verificationResult = new VerificationResult
                {
                    TemplateId = templateId,
                    DocumentName = result.DocumentName,
                    DocumentPath = result.DocumentPath,
                    VerificationDate = result.VerificationDate,
                    TotalMismatches = result.TotalMismatches,
                    Status = result.Status,
                    ErrorMessage = result.ErrorMessage,
                    CreatedBy = result.CreatedBy,
                    CreatedOn = result.CreatedOn
                };

                context.VerificationResults.Add(verificationResult);
                await context.SaveChangesAsync();

                // Add mismatches
                foreach (var mismatch in result.Mismatches)
                {
                    var mismatchEntity = new Mismatch
                    {
                        VerificationResultId = verificationResult.Id,
                        ContextKey = mismatch.ContextKey,
                        Location = mismatch.Location,
                        StructuralRole = mismatch.StructuralRole,
                        Expected = JsonSerializer.Serialize(mismatch.Expected),
                        Actual = JsonSerializer.Serialize(mismatch.Actual),
                        MismatchFields = string.Join(",", mismatch.MismatchFields),
                        SampleText = mismatch.SampleText,
                        Severity = mismatch.Severity,
                        CreatedOn = mismatch.CreatedOn
                    };

                    context.Mismatches.Add(mismatchEntity);
                }

                await context.SaveChangesAsync();
                result.Id = verificationResult.Id;
            }
        }

        private static bool IsStructuralElementEnhanced(TextStyle documentStyle)
        {
            var structuralRoles = new[]
            {
                "TableCell", "TableCellParagraph", "Paragraph", "Heading", "Body", "Table"
            };

            return documentStyle.FormattingContext != null &&
                   structuralRoles.Contains(documentStyle.FormattingContext.StructuralRole);
        }

        private static MismatchReportDto CreateUnexpectedFormattingMismatchEnhanced(TextStyle documentStyle, string contextKey)
        {
            var extraFormatting = GetDirectFormattingPropertiesEnhanced(documentStyle);
            var mismatchedFields = extraFormatting.Keys.Select(k => $"Extra_{k}").ToList();

            return new MismatchReportDto
            {
                ContextKey = contextKey,
                Location = GenerateEnhancedLocation(null, documentStyle.FormattingContext),
                StructuralRole = documentStyle.FormattingContext?.StructuralRole ?? "Unknown",
                Expected = new Dictionary<string, object> { { "ExtraFormatting", "None" } },
                Actual = extraFormatting,
                MismatchedFields = mismatchedFields,
                SampleText = documentStyle.FormattingContext?.SampleText ?? "",
                Severity = "Medium",
                RecommendedAction = $"Remove the extra formatting from this {documentStyle.FormattingContext?.StructuralRole.ToLower() ?? "element"} to match the template"
            };
        }

        /// <summary>
        /// Gets properties that come from direct formatting patterns
        /// </summary>
        private static Dictionary<string, object> GetDirectFormattingPropertiesEnhanced(TextStyle style)
        {
            var properties = new Dictionary<string, object>();

            foreach (var pattern in style.DirectFormatPatterns)
            {
                if (pattern.IsBold.HasValue) properties["IsBold"] = pattern.IsBold.Value;
                if (pattern.IsItalic.HasValue) properties["IsItalic"] = pattern.IsItalic.Value;
                if (pattern.IsUnderline.HasValue) properties["IsUnderline"] = pattern.IsUnderline.Value;
                if (pattern.IsStrikethrough.HasValue) properties["IsStrikethrough"] = pattern.IsStrikethrough.Value;
                if (!string.IsNullOrEmpty(pattern.FontFamily)) properties["FontFamily"] = pattern.FontFamily;
                if (pattern.FontSize.HasValue) properties["FontSize"] = pattern.FontSize.Value;
                if (!string.IsNullOrEmpty(pattern.Color)) properties["Color"] = pattern.Color;
            }

            return properties;
        }
    }
} 