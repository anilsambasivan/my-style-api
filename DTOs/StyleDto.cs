using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.DTOs
{
    public class TextStyleDto
    {
        public int Id { get; set; }
        public int TemplateId { get; set; }
        public string Name { get; set; } = string.Empty;
        public string FontFamily { get; set; } = string.Empty;
        public float FontSize { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public bool IsStrikethrough { get; set; }
        public bool IsAllCaps { get; set; }
        public bool IsSmallCaps { get; set; }
        public bool HasShadow { get; set; }
        public bool HasOutline { get; set; }
        public string Highlighting { get; set; } = string.Empty;
        public float CharacterSpacing { get; set; }
        public string Language { get; set; } = string.Empty;
        public string Color { get; set; } = string.Empty;
        public string Alignment { get; set; } = string.Empty;
        public float SpacingBefore { get; set; }
        public float SpacingAfter { get; set; }
        public float IndentationLeft { get; set; }
        public float IndentationRight { get; set; }
        public float FirstLineIndent { get; set; }
        public bool WidowOrphanControl { get; set; }
        public bool KeepWithNext { get; set; }
        public bool KeepTogether { get; set; }
        public List<TabStopDto> TabStops { get; set; } = new List<TabStopDto>();
        public string BorderStyle { get; set; } = string.Empty;
        public string BorderColor { get; set; } = string.Empty;
        public float BorderWidth { get; set; }
        public string BorderDirections { get; set; } = string.Empty;
        public string ListNumberStyle { get; set; } = string.Empty;
        public int ListLevel { get; set; }
        public float LineSpacing { get; set; }
        public string StyleType { get; set; } = string.Empty;
        public string BasedOnStyle { get; set; } = string.Empty;
        public int Priority { get; set; }
        public string StyleSignature { get; set; } = string.Empty;
        public FormattingContextDto FormattingContext { get; set; } = new FormattingContextDto();
        public List<DirectFormatPatternDto> DirectFormatPatterns { get; set; } = new List<DirectFormatPatternDto>();
        public string CreatedBy { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
        public string ModifiedBy { get; set; } = string.Empty;
        public DateTime? ModifiedOn { get; set; }
        public int Version { get; set; }
    }

    public class CreateTextStyleDto
    {
        [Required]
        public int TemplateId { get; set; }

        [Required]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;

        [Required]
        [StringLength(100)]
        public string FontFamily { get; set; } = string.Empty;

        [Required]
        [Range(1, 144)]
        public float FontSize { get; set; }

        [Required]
        [StringLength(10)]
        public string Color { get; set; } = "#000000";

        [Required]
        [StringLength(50)]
        public string Alignment { get; set; } = "Left";

        [Required]
        [StringLength(50)]
        public string StyleType { get; set; } = "Paragraph";

        [Required]
        [StringLength(100)]
        public string CreatedBy { get; set; } = string.Empty;

        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public float SpacingBefore { get; set; }
        public float SpacingAfter { get; set; }
        public float IndentationLeft { get; set; }
        public float IndentationRight { get; set; }
        public float FirstLineIndent { get; set; }
        public float LineSpacing { get; set; } = 1.0f;
    }

    public class UpdateTextStyleDto
    {
        [Required]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;

        [Required]
        [StringLength(100)]
        public string FontFamily { get; set; } = string.Empty;

        [Required]
        [Range(1, 144)]
        public float FontSize { get; set; }

        [Required]
        [StringLength(10)]
        public string Color { get; set; } = string.Empty;

        [Required]
        [StringLength(50)]
        public string Alignment { get; set; } = string.Empty;

        [Required]
        [StringLength(100)]
        public string ModifiedBy { get; set; } = string.Empty;

        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public bool IsStrikethrough { get; set; }
        public bool IsAllCaps { get; set; }
        public bool IsSmallCaps { get; set; }
        public float SpacingBefore { get; set; }
        public float SpacingAfter { get; set; }
        public float IndentationLeft { get; set; }
        public float IndentationRight { get; set; }
        public float FirstLineIndent { get; set; }
        public float LineSpacing { get; set; }
        public int Priority { get; set; }
    }

    public class DirectFormatPatternDto
    {
        public int Id { get; set; }
        public int TextStyleId { get; set; }
        public string PatternName { get; set; } = string.Empty;
        public string Context { get; set; } = string.Empty;
        public string? FontFamily { get; set; }
        public float? FontSize { get; set; }
        public bool? IsBold { get; set; }
        public bool? IsItalic { get; set; }
        public bool? IsUnderline { get; set; }
        public string? Color { get; set; }
        public string SampleText { get; set; } = string.Empty;
        public int OccurrenceCount { get; set; }
        public string CreatedBy { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
    }

    public class FormattingContextDto
    {
        public int Id { get; set; }
        public string StyleName { get; set; } = string.Empty;
        public string ElementType { get; set; } = string.Empty;
        public string ParentContext { get; set; } = string.Empty;
        public string StructuralRole { get; set; } = string.Empty;
        public string SampleText { get; set; } = string.Empty;
        public string ContentControlTag { get; set; } = string.Empty;
        public string ContentControlType { get; set; } = string.Empty;
        public int SectionIndex { get; set; }
        public int TableIndex { get; set; }
        public int RowIndex { get; set; }
        public int CellIndex { get; set; }
        public int ParagraphIndex { get; set; }
        public int RunIndex { get; set; }
        public string ContextKey { get; set; } = string.Empty;
        public Dictionary<string, string> ContentControlProperties { get; set; } = new Dictionary<string, string>();
    }

    public class TabStopDto
    {
        public int Id { get; set; }
        public float Position { get; set; }
        public string Alignment { get; set; } = string.Empty;
        public string Leader { get; set; } = string.Empty;
    }

    public class StyleComparisonDto
    {
        public TextStyleDto TemplateStyle { get; set; } = new TextStyleDto();
        public TextStyleDto DocumentStyle { get; set; } = new TextStyleDto();
        public List<string> DifferentProperties { get; set; } = new List<string>();
        public bool IsMatch { get; set; }
        public string ComparisonSummary { get; set; } = string.Empty;
    }
} 