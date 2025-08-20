using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DocStyleVerify.API.Models
{
    public class FormattingContext
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        [StringLength(200)]
        public string StyleName { get; set; } = string.Empty;
        
        [Required]
        [StringLength(100)]
        public string ElementType { get; set; } = string.Empty;
        
        [StringLength(500)]
        public string ParentContext { get; set; } = string.Empty;
        
        [StringLength(100)]
        public string StructuralRole { get; set; } = string.Empty;
        
        [StringLength(5000)]
        public string SampleText { get; set; } = string.Empty;
        
        [StringLength(200)]
        public string ContentControlTag { get; set; } = string.Empty;
        
        [StringLength(100)]
        public string ContentControlType { get; set; } = string.Empty;
        
        public int SectionIndex { get; set; } = 0;
        public int TableIndex { get; set; } = -1;
        public int RowIndex { get; set; } = -1;
        public int CellIndex { get; set; } = -1;
        public int ParagraphIndex { get; set; } = -1;
        public int RunIndex { get; set; } = -1;
        
        // Enhanced properties for complex structures
        public int NestedTableLevel { get; set; } = 0;
        public int ListLevel { get; set; } = -1;
        
        [StringLength(100)]
        public string ListNumberStyle { get; set; } = string.Empty;
        
        [StringLength(100)]
        public string HeaderFooterType { get; set; } = string.Empty; // "Header", "Footer", "FirstPageHeader", etc.
        
        public bool IsInHeaderFooter { get; set; } = false;
        
        [StringLength(100)]
        public string DocumentPartType { get; set; } = string.Empty; // "MainDocument", "Header", "Footer", etc.
        
        public int? RowSpan { get; set; }
        public int? ColSpan { get; set; }
        public bool IsMergedCell { get; set; } = false;
        
        [StringLength(50)]
        public string CellMergeType { get; set; } = string.Empty; // "Start", "Continue", "End"
        
        public int ContentControlNestingLevel { get; set; } = 0;
        
        [StringLength(200)]
        public string ContentControlTitle { get; set; } = string.Empty;
        
        [Column(TypeName = "jsonb")]
        public Dictionary<string, string> ContentControlProperties { get; set; } = new Dictionary<string, string>();
        
        // Context key for fast lookup
        [StringLength(500)]
        [Column("context_key")]
        public string ContextKey { get; set; } = string.Empty;
        
        // Navigation property
        public virtual TextStyle? TextStyle { get; set; }
    }
} 