using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DocStyleVerify.API.Models
{
    public class TextStyle
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        public int TemplateId { get; set; }
        
        [Required]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;
        
        [Required]
        [StringLength(100)]
        public string FontFamily { get; set; } = string.Empty;
        
        [Required]
        public float FontSize { get; set; }
        
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public bool IsStrikethrough { get; set; }
        public bool IsAllCaps { get; set; }
        public bool IsSmallCaps { get; set; }
        public bool HasShadow { get; set; }
        public bool HasOutline { get; set; }
        
        [StringLength(50)]
        public string Highlighting { get; set; } = string.Empty;
        
        public float CharacterSpacing { get; set; } = 0;
        
        [StringLength(50)]
        public string Language { get; set; } = string.Empty;
        
        [Required]
        [StringLength(10)]
        public string Color { get; set; } = string.Empty; // Always RGB hex
        
        [Required]
        [StringLength(50)]
        public string Alignment { get; set; } = string.Empty;
        
        public float SpacingBefore { get; set; } = 0;
        public float SpacingAfter { get; set; } = 0;
        public float IndentationLeft { get; set; } = 0;
        public float IndentationRight { get; set; } = 0;
        public float FirstLineIndent { get; set; } = 0;
        public bool WidowOrphanControl { get; set; }
        public bool KeepWithNext { get; set; }
        public bool KeepTogether { get; set; }
        
        [StringLength(50)]
        public string BorderStyle { get; set; } = string.Empty;
        
        [StringLength(10)]
        public string BorderColor { get; set; } = string.Empty;
        
        public float BorderWidth { get; set; } = 0;
        
        [StringLength(100)]
        public string BorderDirections { get; set; } = string.Empty;
        
        [StringLength(50)]
        public string ListNumberStyle { get; set; } = string.Empty;
        
        public int ListLevel { get; set; } = 0;
        public float LineSpacing { get; set; } = 1.0f;
        
        [Required]
        [StringLength(50)]
        public string StyleType { get; set; } = string.Empty; // Paragraph, Character, Table, List, Section
        
        [StringLength(200)]
        public string BasedOnStyle { get; set; } = string.Empty;
        
        public int Priority { get; set; } = 0;
        
        // Table properties
        [StringLength(50)]
        public string TableBorderStyle { get; set; } = string.Empty;
        
        [StringLength(10)]
        public string TableBorderColor { get; set; } = string.Empty;
        
        public float TableBorderWidth { get; set; } = 0;
        
        [StringLength(50)]
        public string TableCellSpacing { get; set; } = string.Empty;
        
        [StringLength(50)]
        public string TableCellPadding { get; set; } = string.Empty;
        
        [StringLength(50)]
        public string TableAlignment { get; set; } = string.Empty;
        
        [StringLength(50)]
        public string TableLeftBorderStyle { get; set; } = string.Empty;
        
        [StringLength(50)]
        public string TableRightBorderStyle { get; set; } = string.Empty;
        
        [StringLength(50)]
        public string TableBottomBorderStyle { get; set; } = string.Empty;
        
        [StringLength(10)]
        public string CellBackgroundColor { get; set; } = string.Empty;
        
        // Enhanced Table properties
        public float TableWidth { get; set; } = 0;
        public float TableCellSpacingValue { get; set; } = 0;
        public string TableShadingColor { get; set; } = string.Empty;
        public string TableShadingPattern { get; set; } = string.Empty;
        public float TableCellWidth { get; set; } = 0;
        public float TableRowHeight { get; set; } = 0;
        public string TableRowHeightRule { get; set; } = string.Empty;
        public bool IsTableHeader { get; set; } = false;
        public bool RepeatOnNewPage { get; set; } = false;
        public float TableCellTopMargin { get; set; } = 0;
        public float TableCellBottomMargin { get; set; } = 0;
        public float TableCellLeftMargin { get; set; } = 0;
        public float TableCellRightMargin { get; set; } = 0;
        public string VerticalAlignment { get; set; } = string.Empty;
        public string TextDirection { get; set; } = string.Empty;
        
        // List/Numbering properties
        public int NumberingId { get; set; } = 0;
        public int AbstractNumberingId { get; set; } = 0;
        public string ListStartValue { get; set; } = string.Empty;
        public string ListNumberFormat { get; set; } = string.Empty;
        public string ListLevelText { get; set; } = string.Empty;
        public string ListBulletFont { get; set; } = string.Empty;
        public string ListBulletChar { get; set; } = string.Empty;
        public float ListTabPosition { get; set; } = 0;
        public float ListIndentPosition { get; set; } = 0;
        
        // Image/Drawing properties
        public float ImageWidth { get; set; } = 0;
        public float ImageHeight { get; set; } = 0;
        public string ImageWrapType { get; set; } = string.Empty;
        public string ImagePosition { get; set; } = string.Empty;
        public float ImageDistanceFromText { get; set; } = 0;
        public string ImageAltText { get; set; } = string.Empty;
        public string ImageTitle { get; set; } = string.Empty;
        public bool ImageLockAspectRatio { get; set; } = false;
        
        // Hyperlink properties
        public string HyperlinkUrl { get; set; } = string.Empty;
        public string HyperlinkTooltip { get; set; } = string.Empty;
        public string HyperlinkTarget { get; set; } = string.Empty;
        public bool HyperlinkVisited { get; set; } = false;
        
        // Field properties
        public string FieldType { get; set; } = string.Empty;
        public string FieldCode { get; set; } = string.Empty;
        public string FieldResult { get; set; } = string.Empty;
        public bool FieldLocked { get; set; } = false;
        public bool FieldDirty { get; set; } = false;
        
        // Section properties
        public string SectionType { get; set; } = string.Empty;
        public int ColumnCount { get; set; } = 1;
        public float ColumnSpacing { get; set; } = 0;
        public bool HasColumnSeparator { get; set; } = false;
        public float HeaderDistance { get; set; } = 0;
        public float FooterDistance { get; set; } = 0;
        public string PageNumberFormat { get; set; } = string.Empty;
        public int PageNumberStart { get; set; } = 1;
        public bool MirrorMargins { get; set; } = false;
        
        // Page/Section properties
        public float PageWidth { get; set; } = 0;
        public float PageHeight { get; set; } = 0;
        public float TopMargin { get; set; } = 0;
        public float BottomMargin { get; set; } = 0;
        public float LeftMargin { get; set; } = 0;
        public float RightMargin { get; set; } = 0;
        
        [StringLength(50)]
        public string Orientation { get; set; } = string.Empty;
        
        [Required]
        [StringLength(500)]
        public string StyleSignature { get; set; } = string.Empty; // for strict fast comparison
        
        // Audit fields
        [Required]
        [StringLength(100)]
        public string CreatedBy { get; set; } = string.Empty;
        
        [Required]
        public DateTime CreatedOn { get; set; } = DateTime.UtcNow;
        
        [StringLength(100)]
        public string ModifiedBy { get; set; } = string.Empty;
        
        public DateTime? ModifiedOn { get; set; }
        
        public int Version { get; set; } = 1;
        
        // Navigation properties
        public virtual ICollection<DirectFormatPattern> DirectFormatPatterns { get; set; } = new List<DirectFormatPattern>();
        public virtual FormattingContext FormattingContext { get; set; } = new FormattingContext();
        public virtual ICollection<TabStop> TabStops { get; set; } = new List<TabStop>();
        public virtual Template Template { get; set; } = null!;
        
        [NotMapped]
        public int FormattingContextId { get; set; }
    }
} 