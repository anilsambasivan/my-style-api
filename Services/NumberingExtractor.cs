using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocStyleVerify.API.Models;
using System.Xml;

namespace DocStyleVerify.API.Services
{
    /// <summary>
    /// Service for extracting numbering information from numbering.xml
    /// </summary>
    public static class NumberingExtractor
    {
        /// <summary>
        /// Extracts numbering definitions from the document's numbering.xml
        /// </summary>
        /// <param name="mainPart">The main document part</param>
        /// <param name="templateId">Template ID</param>
        /// <param name="createdBy">User who initiated the extraction</param>
        /// <returns>List of numbering definitions</returns>
        public static List<NumberingDefinition> ExtractNumberingDefinitions(MainDocumentPart mainPart, int templateId, string createdBy)
        {
            var numberingDefinitions = new List<NumberingDefinition>();
            
            var numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart?.Numbering == null)
                return numberingDefinitions;

            // Extract abstract numbering definitions
            var abstractNums = numberingPart.Numbering.Descendants<AbstractNum>().ToList();
            var numberingInstances = numberingPart.Numbering.Descendants<NumberingInstance>().ToList();

            foreach (var abstractNum in abstractNums)
            {
                var numberingDef = CreateNumberingDefinitionFromAbstractNum(abstractNum, templateId, createdBy);
                if (numberingDef != null)
                {
                    // Find corresponding numbering instances
                    var relatedInstances = numberingInstances.Where(ni => 
                        ni.AbstractNumId?.Val?.Value == abstractNum.AbstractNumberId?.Value).ToList();
                    
                    foreach (var instance in relatedInstances)
                    {
                        if (instance.NumberID?.Value != null)
                        {
                            numberingDef.NumberingId = instance.NumberID.Value;
                            break; // Use the first matching instance
                        }
                    }

                    numberingDefinitions.Add(numberingDef);
                }
            }

            return numberingDefinitions;
        }

        /// <summary>
        /// Creates a NumberingDefinition entity from an AbstractNum element
        /// </summary>
        private static NumberingDefinition? CreateNumberingDefinitionFromAbstractNum(AbstractNum abstractNum, int templateId, string createdBy)
        {
            if (abstractNum.AbstractNumberId?.Value == null)
                return null;

            var numberingDef = new NumberingDefinition
            {
                TemplateId = templateId,
                AbstractNumId = abstractNum.AbstractNumberId.Value,
                NumberingId = 0, // Will be set later from NumberingInstance
                Name = GetNumberingName(abstractNum),
                Type = GetNumberingType(abstractNum),
                CreatedBy = createdBy,
                CreatedOn = DateTime.UtcNow,
                RawXml = abstractNum.OuterXml
            };

            // Extract numbering levels
            var levels = abstractNum.Descendants<Level>().ToList();
            foreach (var level in levels)
            {
                var numberingLevel = CreateNumberingLevelFromLevel(level, numberingDef.Id);
                if (numberingLevel != null)
                {
                    numberingDef.NumberingLevels.Add(numberingLevel);
                }
            }

            return numberingDef;
        }

        /// <summary>
        /// Creates a NumberingLevel entity from a Level element
        /// </summary>
        private static NumberingLevel? CreateNumberingLevelFromLevel(Level level, int numberingDefinitionId)
        {
            if (level.LevelIndex?.Value == null)
                return null;

            var numberingLevel = new NumberingLevel
            {
                NumberingDefinitionId = numberingDefinitionId,
                Level = level.LevelIndex.Value,
                NumberFormat = level.NumberingFormat?.Val?.ToString() ?? string.Empty,
                LevelText = level.LevelText?.Val?.Value ?? string.Empty,
                LevelJustification = level.LevelJustification?.Val?.ToString() ?? string.Empty,
                StartValue = level.StartNumberingValue?.Val?.Value ?? 1,
                IsLegal = level.IsLegalNumberingStyle?.Val?.Value ?? false,
                RawXml = level.OuterXml
            };

            // Extract font properties from run properties
            var runProps = level.GetFirstChild<RunProperties>();
            if (runProps != null)
            {
                ExtractFontProperties(runProps, numberingLevel);
            }

            // Extract paragraph properties
            var paraProps = level.GetFirstChild<ParagraphProperties>();
            if (paraProps != null)
            {
                ExtractParagraphProperties(paraProps, numberingLevel);
            }

            return numberingLevel;
        }

        /// <summary>
        /// Extracts font properties from RunProperties
        /// </summary>
        private static void ExtractFontProperties(RunProperties runProps, NumberingLevel numberingLevel)
        {
            if (runProps.RunFonts != null)
            {
                numberingLevel.FontFamily = runProps.RunFonts.Ascii?.Value ?? 
                                           runProps.RunFonts.HighAnsi?.Value ?? 
                                           runProps.RunFonts.ComplexScript?.Value ?? 
                                           runProps.RunFonts.EastAsia?.Value ?? 
                                           string.Empty;
            }

            if (runProps.FontSize != null && float.TryParse(runProps.FontSize.Val?.Value, out var fontSize))
            {
                numberingLevel.FontSize = fontSize / 2f; // Convert half-points to points
            }

            numberingLevel.IsBold = runProps.Bold != null && (runProps.Bold.Val?.Value ?? true);
            numberingLevel.IsItalic = runProps.Italic != null && (runProps.Italic.Val?.Value ?? true);

            if (runProps.Color != null)
            {
                numberingLevel.Color = runProps.Color.Val?.Value ?? string.Empty;
            }
        }

        /// <summary>
        /// Extracts paragraph properties from ParagraphProperties
        /// </summary>
        private static void ExtractParagraphProperties(ParagraphProperties paraProps, NumberingLevel numberingLevel)
        {
            if (paraProps.Indentation != null)
            {
                if (paraProps.Indentation.Left != null && 
                    int.TryParse(paraProps.Indentation.Left.Value, out var leftIndent))
                {
                    numberingLevel.IndentationLeft = ConvertTwipsToPoints(leftIndent);
                }

                if (paraProps.Indentation.Hanging != null && 
                    int.TryParse(paraProps.Indentation.Hanging.Value, out var hangingIndent))
                {
                    numberingLevel.IndentationHanging = ConvertTwipsToPoints(hangingIndent);
                }
            }

            // Extract tab stops for numbering
            if (paraProps.Tabs != null)
            {
                var firstTab = paraProps.Tabs.Descendants<DocumentFormat.OpenXml.Wordprocessing.TabStop>().FirstOrDefault();
                if (firstTab?.Position != null && 
                    int.TryParse(firstTab.Position.Value.ToString(), out var tabPosition))
                {
                    numberingLevel.TabStopPosition = ConvertTwipsToPoints(tabPosition);
                }
            }
        }

        /// <summary>
        /// Gets the numbering name from AbstractNum
        /// </summary>
        private static string GetNumberingName(AbstractNum abstractNum)
        {
            // Try to get name from various sources
            var name = abstractNum.AbstractNumDefinitionName?.Val?.Value;
            if (!string.IsNullOrEmpty(name))
                return name;

            var styleLink = abstractNum.StyleLink?.Val?.Value;
            if (!string.IsNullOrEmpty(styleLink))
                return styleLink;

            // NumStyleLink property doesn't exist, skip this check

            return $"Numbering_{abstractNum.AbstractNumberId?.Value ?? 0}";
        }

        /// <summary>
        /// Determines the numbering type from AbstractNum
        /// </summary>
        private static string GetNumberingType(AbstractNum abstractNum)
        {
            var levels = abstractNum.Descendants<Level>().ToList();
            if (!levels.Any())
                return "Unknown";

            var firstLevel = levels.First();
            var numberFormat = firstLevel.NumberingFormat?.Val?.ToString();

            return numberFormat switch
            {
                "bullet" => "Bullet",
                "decimal" => "Number",
                "lowerLetter" => "LowerLetter",
                "upperLetter" => "UpperLetter",
                "lowerRoman" => "LowerRoman",
                "upperRoman" => "UpperRoman",
                "decimalZero" => "DecimalZero",
                "ordinal" => "Ordinal",
                "cardinalText" => "CardinalText",
                "ordinalText" => "OrdinalText",
                "hex" => "Hex",
                "chicago" => "Chicago",
                "ideographDigital" => "IdeographDigital",
                "japaneseCounting" => "JapaneseCounting",
                "aiueo" => "Aiueo",
                "iroha" => "Iroha",
                "decimalFullWidth" => "DecimalFullWidth",
                "decimalHalfWidth" => "DecimalHalfWidth",
                "japaneseLegal" => "JapaneseLegal",
                "japaneseDigitalTenThousand" => "JapaneseDigitalTenThousand",
                "decimalEnclosedCircle" => "DecimalEnclosedCircle",
                "decimalFullWidth2" => "DecimalFullWidth2",
                "aiueoFullWidth" => "AiueoFullWidth",
                "irohaFullWidth" => "IrohaFullWidth",
                "decimalZeroFullWidth" => "DecimalZeroFullWidth",
                "bulletFullWidth" => "BulletFullWidth",
                "ganada" => "Ganada",
                "chosung" => "Chosung",
                "decimalEnclosedFullstop" => "DecimalEnclosedFullstop",
                "decimalEnclosedParen" => "DecimalEnclosedParen",
                "decimalEnclosedCircleChinese" => "DecimalEnclosedCircleChinese",
                "ideographEnclosedCircle" => "IdeographEnclosedCircle",
                "ideographTraditional" => "IdeographTraditional",
                "ideographZodiac" => "IdeographZodiac",
                "ideographZodiacTraditional" => "IdeographZodiacTraditional",
                "taiwaneseCounting" => "TaiwaneseCounting",
                "ideographLegalTraditional" => "IdeographLegalTraditional",
                "taiwaneseCountingThousand" => "TaiwaneseCountingThousand",
                "taiwaneseDigital" => "TaiwaneseDigital",
                "chineseCounting" => "ChineseCounting",
                "chineseLegalSimplified" => "ChineseLegalSimplified",
                "chineseCountingThousand" => "ChineseCountingThousand",
                "koreanDigital" => "KoreanDigital",
                "koreanCounting" => "KoreanCounting",
                "koreanLegal" => "KoreanLegal",
                "koreanDigital2" => "KoreanDigital2",
                "vietnameseCounting" => "VietnameseCounting",
                "russianLower" => "RussianLower",
                "russianUpper" => "RussianUpper",
                "none" => "None",
                "numberInDash" => "NumberInDash",
                "hebrew1" => "Hebrew1",
                "arabic1" => "Arabic1",
                "hebrew2" => "Hebrew2",
                "arabic2" => "Arabic2",
                "hindiLetter1" => "HindiLetter1",
                "hindiLetter2" => "HindiLetter2",
                "hindiArabic" => "HindiArabic",
                "hindiCounting" => "HindiCounting",
                "thaiLetter" => "ThaiLetter",
                "thaiArabic" => "ThaiArabic",
                "thaiCounting" => "ThaiCounting",
                "bahtText" => "BahtText",
                "dollarText" => "DollarText",
                "custom" => "Custom",
                _ => "Unknown"
            };
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