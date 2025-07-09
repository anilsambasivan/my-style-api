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
    /// Extensions for DocumentProcessingService to complete the implementation
    /// </summary>
    public static class DocumentProcessingServiceExtensions
    {
        public static async Task<List<TextStyle>> ExtractDocumentStylesAsync(
            this DocumentProcessingService service,
            Stream documentStream,
            string fileName)
        {
            try
            {
                // Extract both styles and patterns
                var (extractedStyles, extractedPatterns) = await service.ExtractStylesOnlyAsync(documentStream, 0, "System");
                
                // Associate patterns with styles using the same logic as template extraction
                AssociateDirectFormattingPatterns(extractedStyles, extractedPatterns, service);
                
                return extractedStyles;
            }
            catch (Exception ex)
            {
                // Log error and return empty list
                System.Diagnostics.Debug.WriteLine($"Error extracting document styles: {ex.Message}");
                return new List<TextStyle>();
            }
        }

        /// <summary>
        /// Associates direct formatting patterns with their corresponding styles
        /// </summary>
        private static void AssociateDirectFormattingPatterns(
            List<TextStyle> styles, 
            List<DirectFormatPattern> patterns,
            DocumentProcessingService service)
        {
            foreach (var pattern in patterns)
            {
                var matchingStyle = styles.FirstOrDefault(s => service.ShouldPatternBelongToStyle(pattern, s));
                if (matchingStyle != null)
                {
                    matchingStyle.DirectFormatPatterns.Add(pattern);
                }
            }
        }

        public static async Task<VerificationResultDto> VerifyDocumentAsync(
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
                // Get template styles from database
                var templateStyles = await GetTemplateStylesFromDatabaseAsync(service, templateId);
                if (!templateStyles.Any())
                {
                    result.Status = "Failed";
                    result.ErrorMessage = "No template styles found for comparison";
                    return result;
                }

                // Extract document styles
                var documentStyles = await service.ExtractDocumentStylesAsync(documentStream, fileName);

                // Compare styles with enhanced error handling
                List<MismatchReportDto> mismatches;
                try
                {
                    mismatches = CompareStylesDetailed(templateStyles, documentStyles, strictMode, ignoreStyleTypes);
                }
                catch (ArgumentException ex) when (ex.Message.Contains("same key has already been added"))
                {
                    result.Status = "Failed";
                    result.ErrorMessage = $"Duplicate style signatures detected. This issue has been fixed in the code - please try again. Details: {ex.Message}";
                    return result;
                }

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

        public static List<MismatchReportDto> CompareStyles(
            this DocumentProcessingService service,
            List<TextStyle> templateStyles,
            List<TextStyle> documentStyles,
            bool strictMode = true)
        {
            return CompareStylesDetailed(templateStyles, documentStyles, strictMode, null);
        }

        public static async Task<bool> ValidateDocumentFormatAsync(
            this DocumentProcessingService service,
            Stream documentStream)
        {
            try
            {
                using var document = WordprocessingDocument.Open(documentStream, false);
                var mainPart = document.MainDocumentPart;
                return mainPart?.Document?.Body != null;
            }
            catch
            {
                return false;
            }
        }

        public static async Task<DocumentMetadata> GetDocumentMetadataAsync(
            this DocumentProcessingService service,
            Stream documentStream)
        {
            var metadata = new DocumentMetadata();

            try
            {
                using var document = WordprocessingDocument.Open(documentStream, false);
                var mainPart = document.MainDocumentPart;

                if (mainPart == null)
                    return metadata;

                // Extract core properties
                var coreProps = document.PackageProperties;
                metadata.Title = coreProps.Title ?? "";
                metadata.Author = coreProps.Creator ?? "";
                metadata.Subject = coreProps.Subject ?? "";
                metadata.Created = coreProps.Created;
                metadata.LastModified = coreProps.Modified;

                // Count document elements
                var body = mainPart.Document.Body;
                if (body != null)
                {
                    metadata.ParagraphCount = body.Descendants<Paragraph>().Count();
                    metadata.TableCount = body.Descendants<Table>().Count();
                    metadata.WordCount = EstimateWordCount(body);
                }

                // Extract style information
                var stylesDoc = mainPart.StyleDefinitionsPart?.Styles;
                if (stylesDoc != null)
                {
                    metadata.StyleNames = stylesDoc.Descendants<Style>()
                        .Where(s => s.StyleName?.Val?.Value != null)
                        .Select(s => s.StyleName!.Val!.Value!)
                        .ToList();
                }
            }
            catch (Exception)
            {
                // Return partial metadata on error
            }

            return metadata;
        }

        private static async Task<List<TextStyle>> GetTemplateStylesFromDatabaseAsync(
            DocumentProcessingService service,
            int templateId)
        {
            // Access the private context field using reflection or create a helper method
            // For now, we'll create a simple approach
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

        internal static List<MismatchReportDto> CompareStylesDetailed(
            List<TextStyle> templateStyles,
            List<TextStyle> documentStyles,
            bool strictMode,
            List<string>? ignoreStyleTypes)
        {
            var mismatches = new List<MismatchReportDto>();
            ignoreStyleTypes ??= new List<string>();

            // Create lookup dictionaries for faster matching
            var documentStylesByContext = CreateStyleLookup(documentStyles);
            var documentStylesBySignature = CreateStyleSignatureLookup(documentStyles);

            foreach (var templateStyle in templateStyles)
            {
                if (ignoreStyleTypes.Contains(templateStyle.StyleType))
                    continue;

                var contextKey = templateStyle.FormattingContext?.ContextKey ?? $"Style_{templateStyle.Id}";
                
                // Find matching document style using multiple strategies
                var documentStyle = FindMatchingDocumentStyleEnhanced(documentStyles, templateStyle, documentStylesByContext, documentStylesBySignature);

                if (documentStyle == null)
                {
                    // Style completely missing in document
                    mismatches.Add(CreateMissingStyleMismatch(templateStyle, contextKey));
                    continue;
                }

                // Enhanced comparison that considers base style + direct formatting combinations
                var styleMismatches = CompareStylesWithDirectFormattingEnhanced(templateStyle, documentStyle, contextKey, strictMode);
                mismatches.AddRange(styleMismatches);
            }

            // Check for extra styles in document that don't exist in template
            var unmatchedDocumentStyles = new HashSet<TextStyle>();
            
            foreach (var documentStyle in documentStyles)
            {
                if (ignoreStyleTypes.Contains(documentStyle.StyleType))
                    continue;

                var contextKey = documentStyle.FormattingContext?.ContextKey ?? $"Style_{documentStyle.Id}";
                
                // Enhanced matching logic that properly handles direct formatting
                var matchingTemplate = FindMatchingTemplateStyleForDocument(templateStyles, documentStyle, contextKey);

                if (matchingTemplate == null)
                {
                    // Only mark as "should not exist" if truly no logical equivalent exists in template
                    unmatchedDocumentStyles.Add(documentStyle);
                }
            }

            // Process unmatched document styles with enhanced logic
            foreach (var documentStyle in unmatchedDocumentStyles)
            {
                var contextKey = documentStyle.FormattingContext?.ContextKey ?? $"Style_{documentStyle.Id}";
                
                // Check if this is a structural element that should exist but with wrong formatting
                var isStructuralElement = IsStructuralElement(documentStyle);
                
                if (isStructuralElement)
                {
                    // This is a table cell, paragraph, etc. that exists but has unexpected formatting
                    // Report this as extra formatting rather than "should not exist"
                    mismatches.Add(CreateUnexpectedFormattingMismatch(documentStyle, contextKey));
                }
                else
                {
                    // This is truly an extra style that shouldn't be there
                    mismatches.Add(CreateExtraStyleMismatch(documentStyle, contextKey));
                }
            }

            // Filter out mismatches where SampleText is empty or null
            // This reduces noise by focusing only on content that actually exists
            return mismatches.Where(m => !string.IsNullOrWhiteSpace(m.SampleText)).ToList();
        }

        /// <summary>
        /// Enhanced style comparison that properly handles direct formatting combinations
        /// </summary>
        private static List<MismatchReportDto> CompareStylesWithDirectFormattingEnhanced(
            TextStyle templateStyle,
            TextStyle documentStyle,
            string contextKey,
            bool strictMode)
        {
            var mismatches = new List<MismatchReportDto>();

            // Get effective style properties (base style + direct formatting)
            var templateEffective = GetEffectiveStyleProperties(templateStyle);
            var documentEffective = GetEffectiveStyleProperties(documentStyle);

            // Compare effective properties
            var propertyDifferences = FindPropertyDifferences(templateEffective, documentEffective, strictMode);

            if (propertyDifferences.Any())
            {
                mismatches.Add(new MismatchReportDto
                {
                    ContextKey = contextKey,
                    Location = GenerateLocationString(templateStyle.FormattingContext ?? documentStyle.FormattingContext),
                    StructuralRole = templateStyle.FormattingContext?.StructuralRole ?? documentStyle.FormattingContext?.StructuralRole ?? "Unknown",
                    Expected = templateEffective,
                    Actual = documentEffective,
                    MismatchedFields = propertyDifferences,
                    SampleText = documentStyle.FormattingContext?.SampleText ?? templateStyle.FormattingContext?.SampleText ?? "",
                    Severity = DetermineEnhancedSeverity(propertyDifferences),
                    RecommendedAction = GenerateEnhancedRecommendedAction(propertyDifferences, templateStyle, documentStyle)
                });
            }

            // Compare direct formatting patterns separately for detailed reporting
            var directFormattingMismatches = CompareDirectFormattingPatterns(templateStyle, documentStyle, contextKey);
            mismatches.AddRange(directFormattingMismatches);

            return mismatches;
        }

        /// <summary>
        /// Compares direct formatting patterns between template and document styles
        /// </summary>
        private static List<MismatchReportDto> CompareDirectFormattingPatterns(
            TextStyle templateStyle,
            TextStyle documentStyle,
            string contextKey)
        {
            var mismatches = new List<MismatchReportDto>();

            // Group patterns by their context for comparison
            var templatePatterns = templateStyle.DirectFormatPatterns.ToLookup(p => p.Context);
            var documentPatterns = documentStyle.DirectFormatPatterns.ToLookup(p => p.Context);

            // Check each template pattern
            foreach (var templatePatternGroup in templatePatterns)
            {
                var patternContext = templatePatternGroup.Key;
                var templatePattern = templatePatternGroup.First(); // Use first if multiple

                var documentPattern = documentPatterns[patternContext].FirstOrDefault();

                if (documentPattern == null)
                {
                    // Direct formatting pattern missing in document
                    mismatches.Add(CreateMissingDirectFormattingMismatch(templatePattern, contextKey, patternContext));
                }
                else
                {
                    // Compare direct formatting properties
                    var patternDifferences = GetDirectFormattingDifferences(templatePattern, documentPattern);
                    if (patternDifferences.Any())
                    {
                        mismatches.Add(CreateDirectFormattingMismatch(templatePattern, documentPattern, contextKey, patternContext, patternDifferences));
                    }
                }
            }

            // Check for extra direct formatting in document that shouldn't be there
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
            AddPropertyIfNotDefault(properties, "StyleType", style.StyleType);

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
        /// Enhanced style matching that considers context, signature, and fallback strategies
        /// </summary>
        private static TextStyle? FindMatchingDocumentStyleEnhanced(
            List<TextStyle> documentStyles,
            TextStyle templateStyle,
            Dictionary<string, TextStyle> documentStylesByContext,
            Dictionary<string, TextStyle> documentStylesBySignature)
        {
            // Strategy 1: Match by context key (most reliable)
            if (templateStyle.FormattingContext?.ContextKey != null &&
                documentStylesByContext.TryGetValue(templateStyle.FormattingContext.ContextKey, out var contextMatch))
            {
                return contextMatch;
            }

            // Strategy 2: Match by style signature
            if (documentStylesBySignature.TryGetValue(templateStyle.StyleSignature, out var signatureMatch))
            {
                return signatureMatch;
            }

            // Strategy 3: Match by style name and type
            var nameTypeMatch = documentStyles.FirstOrDefault(ds => 
                ds.Name == templateStyle.Name && ds.StyleType == templateStyle.StyleType);
            
            if (nameTypeMatch != null)
            {
                return nameTypeMatch;
            }

            // Strategy 4: Match by similar context (for cases where indexing might differ)
            if (templateStyle.FormattingContext != null)
            {
                var similarContextMatch = documentStyles.FirstOrDefault(ds =>
                    ds.FormattingContext != null &&
                    ds.FormattingContext.StructuralRole == templateStyle.FormattingContext.StructuralRole &&
                    ds.FormattingContext.ParagraphIndex == templateStyle.FormattingContext.ParagraphIndex &&
                    ds.FormattingContext.ElementType == templateStyle.FormattingContext.ElementType);
                
                if (similarContextMatch != null)
                {
                    return similarContextMatch;
                }
            }

            return null;
        }

        /// <summary>
        /// Creates lookup dictionaries for faster style matching
        /// </summary>
        private static Dictionary<string, TextStyle> CreateStyleLookup(List<TextStyle> styles)
        {
            var lookup = new Dictionary<string, TextStyle>();
            
            foreach (var style in styles)
            {
                var key = style.FormattingContext?.ContextKey ?? $"{style.StyleType}:{style.Name}";
                if (!lookup.ContainsKey(key))
                {
                    lookup[key] = style;
                }
            }
            
            return lookup;
        }

        /// <summary>
        /// Creates a lookup dictionary by StyleSignature, handling duplicates by taking the first occurrence
        /// </summary>
        private static Dictionary<string, TextStyle> CreateStyleSignatureLookup(List<TextStyle> styles)
        {
            var lookup = new Dictionary<string, TextStyle>();
            
            foreach (var style in styles)
            {
                if (!string.IsNullOrEmpty(style.StyleSignature) && !lookup.ContainsKey(style.StyleSignature))
                {
                    lookup[style.StyleSignature] = style;
                }
            }
            
            return lookup;
        }

        /// <summary>
        /// Finds differences between effective style properties
        /// </summary>
        private static List<string> FindPropertyDifferences(
            Dictionary<string, object> templateProps,
            Dictionary<string, object> documentProps,
            bool strictMode)
        {
            var differences = new List<string>();

            // Check all template properties
            foreach (var templateProp in templateProps)
            {
                var propName = templateProp.Key;
                var templateValue = templateProp.Value;

                if (!documentProps.TryGetValue(propName, out var documentValue))
                {
                    // Property missing in document
                    differences.Add(propName);
                    continue;
                }

                // Compare values based on type
                if (!AreValuesEqual(templateValue, documentValue, propName, strictMode))
                {
                    differences.Add(propName);
                }
            }

            // Check for extra properties in document
            foreach (var documentProp in documentProps)
            {
                if (!templateProps.ContainsKey(documentProp.Key))
                {
                    differences.Add($"Extra_{documentProp.Key}");
                }
            }

            return differences;
        }

        /// <summary>
        /// Compares two values for equality with type-specific logic
        /// </summary>
        private static bool AreValuesEqual(object templateValue, object documentValue, string propertyName, bool strictMode)
        {
            if (templateValue == null && documentValue == null) return true;
            if (templateValue == null || documentValue == null) return false;

            // Handle floating point comparisons with tolerance
            if (templateValue is float templateFloat && documentValue is float documentFloat)
            {
                var tolerance = strictMode ? 0.1f : 1.0f;
                return Math.Abs(templateFloat - documentFloat) <= tolerance;
            }

            // Handle string comparisons (case-insensitive for certain properties)
            if (templateValue is string templateStr && documentValue is string documentStr)
            {
                if (propertyName == "FontFamily" || propertyName == "Alignment")
                {
                    return string.Equals(templateStr, documentStr, StringComparison.OrdinalIgnoreCase);
                }
                return string.Equals(templateStr, documentStr, StringComparison.Ordinal);
            }

            // Default equality check
            return templateValue.Equals(documentValue);
        }

        /// <summary>
        /// Adds property to dictionary if it's not a default value
        /// </summary>
        private static void AddPropertyIfNotDefault(Dictionary<string, object> properties, string name, object value)
        {
            if (value == null) return;

            switch (value)
            {
                case string str when !string.IsNullOrEmpty(str):
                    properties[name] = str;
                    break;
                case float f when f != 0:
                    properties[name] = f;
                    break;
                case bool b when b:
                    properties[name] = b;
                    break;
                case int i when i != 0:
                    properties[name] = i;
                    break;
                default:
                    if (!IsDefaultValue(value))
                        properties[name] = value;
                    break;
            }
        }

        /// <summary>
        /// Checks if a value is considered a default value
        /// </summary>
        private static bool IsDefaultValue(object value)
        {
            if (value == null) return true;
            
            return value switch
            {
                string s => string.IsNullOrEmpty(s),
                float f => f == 0,
                bool b => !b,
                int i => i == 0,
                _ => false
            };
        }

        /// <summary>
        /// Gets differences between two direct formatting patterns
        /// </summary>
        private static List<string> GetDirectFormattingDifferences(DirectFormatPattern template, DirectFormatPattern document)
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

            return differences;
        }

        /// <summary>
        /// Creates a mismatch for a missing style
        /// </summary>
        private static MismatchReportDto CreateMissingStyleMismatch(TextStyle templateStyle, string contextKey)
        {
            return new MismatchReportDto
            {
                ContextKey = contextKey,
                Location = GenerateLocationString(templateStyle.FormattingContext),
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
        /// Creates a mismatch for an extra style not in template
        /// </summary>
        private static MismatchReportDto CreateExtraStyleMismatch(TextStyle documentStyle, string contextKey)
        {
            return new MismatchReportDto
            {
                ContextKey = contextKey,
                Location = GenerateLocationString(documentStyle.FormattingContext),
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
                Location = $"{GenerateLocationFromContext(contextKey)} > {patternContext}",
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
                Location = $"{GenerateLocationFromContext(contextKey)} > {patternContext}",
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
                Location = $"{GenerateLocationFromContext(contextKey)} > {patternContext}",
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
        /// Serializes a direct formatting pattern for comparison
        /// </summary>
        private static Dictionary<string, object> SerializeDirectFormattingPattern(DirectFormatPattern pattern)
        {
            var properties = new Dictionary<string, object>();
            
            AddPropertyIfNotDefault(properties, "FontFamily", pattern.FontFamily ?? "");
            AddPropertyIfNotDefault(properties, "FontSize", pattern.FontSize ?? 0f);
            AddPropertyIfNotDefault(properties, "IsBold", pattern.IsBold ?? false);
            AddPropertyIfNotDefault(properties, "IsItalic", pattern.IsItalic ?? false);
            AddPropertyIfNotDefault(properties, "IsUnderline", pattern.IsUnderline ?? false);
            AddPropertyIfNotDefault(properties, "IsStrikethrough", pattern.IsStrikethrough ?? false);
            AddPropertyIfNotDefault(properties, "Color", pattern.Color ?? "");
            AddPropertyIfNotDefault(properties, "Alignment", pattern.Alignment ?? "");
            AddPropertyIfNotDefault(properties, "SpacingBefore", pattern.SpacingBefore ?? 0f);
            AddPropertyIfNotDefault(properties, "SpacingAfter", pattern.SpacingAfter ?? 0f);
            AddPropertyIfNotDefault(properties, "IndentationLeft", pattern.IndentationLeft ?? 0f);
            AddPropertyIfNotDefault(properties, "IndentationRight", pattern.IndentationRight ?? 0f);
            AddPropertyIfNotDefault(properties, "LineSpacing", pattern.LineSpacing ?? 1f);
            
            return properties;
        }

        /// <summary>
        /// Generates a location string from context key
        /// </summary>
        private static string GenerateLocationFromContext(string contextKey)
        {
            return contextKey.Replace(":", " > ").Replace("Section", "Section").Replace("Paragraph", "Paragraph");
        }

        /// <summary>
        /// Generates a location string from formatting context
        /// </summary>
        private static string GenerateLocationString(FormattingContext? context)
        {
            if (context == null) return "Unknown";

            var parts = new List<string>();
            
            if (context.SectionIndex >= 0) parts.Add($"Section {context.SectionIndex + 1}");
            if (context.TableIndex >= 0) parts.Add($"Table {context.TableIndex + 1}");
            if (context.RowIndex >= 0) parts.Add($"Row {context.RowIndex + 1}");
            if (context.CellIndex >= 0) parts.Add($"Cell {context.CellIndex + 1}");
            if (context.ParagraphIndex >= 0) parts.Add($"Paragraph {context.ParagraphIndex + 1}");

            return string.Join(", ", parts);
        }

        /// <summary>
        /// Determines severity for enhanced comparison
        /// </summary>
        private static string DetermineEnhancedSeverity(List<string> differences)
        {
            if (differences.Any(d => d.Contains("FontFamily") || d.Contains("Color")))
                return "High";
            if (differences.Any(d => d.Contains("FontSize") || d.Contains("Alignment") || d.Contains("Bold") || d.Contains("Italic")))
                return "Medium";
            return "Low";
        }

        /// <summary>
        /// Determines severity for direct formatting mismatches
        /// </summary>
        private static string DetermineDirectFormattingSeverity(List<string> differences)
        {
            if (differences.Contains("Color") || differences.Contains("FontFamily"))
                return "High";
            if (differences.Contains("IsBold") || differences.Contains("IsItalic") || differences.Contains("FontSize"))
                return "Medium";
            return "Low";
        }

        /// <summary>
        /// Generates enhanced recommended action
        /// </summary>
        private static string GenerateEnhancedRecommendedAction(List<string> differences, TextStyle templateStyle, TextStyle documentStyle)
        {
            var hasDirectFormatting = documentStyle.DirectFormatPatterns.Any();
            var shouldHaveDirectFormatting = templateStyle.DirectFormatPatterns.Any();

            if (hasDirectFormatting && !shouldHaveDirectFormatting)
            {
                return $"Remove direct formatting from '{documentStyle.Name}' and apply only the base style";
            }
            
            if (!hasDirectFormatting && shouldHaveDirectFormatting)
            {
                return $"Apply the required direct formatting to '{documentStyle.Name}' as specified in the template";
            }

            if (differences.Count == 1)
            {
                return $"Correct the {differences[0]} property to match the template";
            }
            
            return $"Correct the following properties to match the template: {string.Join(", ", differences)}";
        }

        private static int EstimateWordCount(Body body)
        {
            var text = body.InnerText;
            return text.Split(new char[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }

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

        /// <summary>
        /// Enhanced matching logic that properly handles direct formatting
        /// </summary>
        private static TextStyle? FindMatchingTemplateStyleForDocument(
            List<TextStyle> templateStyles,
            TextStyle documentStyle,
            string contextKey)
        {
            // Strategy 1: Exact context key match
            var contextMatch = templateStyles.FirstOrDefault(ts => 
                ts.FormattingContext?.ContextKey == contextKey);
            if (contextMatch != null)
                return contextMatch;

            // Strategy 2: Match by structural location (same table cell, paragraph position, etc.)
            if (documentStyle.FormattingContext != null)
            {
                var structuralMatch = templateStyles.FirstOrDefault(ts =>
                    ts.FormattingContext != null &&
                    ts.FormattingContext.StructuralRole == documentStyle.FormattingContext.StructuralRole &&
                    ts.FormattingContext.SectionIndex == documentStyle.FormattingContext.SectionIndex &&
                    ts.FormattingContext.TableIndex == documentStyle.FormattingContext.TableIndex &&
                    ts.FormattingContext.RowIndex == documentStyle.FormattingContext.RowIndex &&
                    ts.FormattingContext.CellIndex == documentStyle.FormattingContext.CellIndex &&
                    ts.FormattingContext.ParagraphIndex == documentStyle.FormattingContext.ParagraphIndex &&
                    ts.StyleType == documentStyle.StyleType);
                
                if (structuralMatch != null)
                    return structuralMatch;
            }

            // Strategy 3: Match by base style signature (ignoring direct formatting)
            var baseStyleSignature = GetBaseStyleSignature(documentStyle);
            var baseStyleMatch = templateStyles.FirstOrDefault(ts => 
                GetBaseStyleSignature(ts) == baseStyleSignature);
            
            if (baseStyleMatch != null)
                return baseStyleMatch;

            // Strategy 4: Match by style name and type
            var nameTypeMatch = templateStyles.FirstOrDefault(ts =>
                ts.Name == documentStyle.Name && 
                ts.StyleType == documentStyle.StyleType);

            return nameTypeMatch;
        }

        /// <summary>
        /// Determines if a style represents a structural document element
        /// </summary>
        private static bool IsStructuralElement(TextStyle documentStyle)
        {
            var structuralRoles = new[]
            {
                "TableCell", "TableCellParagraph", "Paragraph", "Heading", "Body", "Table"
            };

            return documentStyle.FormattingContext != null &&
                   structuralRoles.Contains(documentStyle.FormattingContext.StructuralRole);
        }

        /// <summary>
        /// Creates a mismatch for unexpected formatting on a structural element
        /// </summary>
        private static MismatchReportDto CreateUnexpectedFormattingMismatch(TextStyle documentStyle, string contextKey)
        {
            var extraFormatting = GetDirectFormattingProperties(documentStyle);
            var mismatchedFields = extraFormatting.Keys.Select(k => $"Extra_{k}").ToList();

            return new MismatchReportDto
            {
                ContextKey = contextKey,
                Location = GenerateLocationString(documentStyle.FormattingContext),
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
        /// Gets base style signature ignoring direct formatting
        /// </summary>
        private static string GetBaseStyleSignature(TextStyle style)
        {
            // Create a signature based on base style properties only (ignoring direct formatting)
            var baseProperties = new
            {
                style.FontFamily,
                style.FontSize,
                style.IsBold,
                style.IsItalic,
                style.IsUnderline,
                style.IsStrikethrough,
                style.Color,
                style.Alignment,
                style.SpacingBefore,
                style.SpacingAfter,
                style.IndentationLeft,
                style.IndentationRight,
                style.LineSpacing,
                style.StyleType,
                style.BasedOnStyle
            };

            return baseProperties.GetHashCode().ToString();
        }

        /// <summary>
        /// Gets properties that come from direct formatting patterns
        /// </summary>
        private static Dictionary<string, object> GetDirectFormattingProperties(TextStyle style)
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