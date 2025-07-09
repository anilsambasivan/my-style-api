using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DocStyleVerify.API.Models
{
    public class DirectFormatPattern
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        public int TextStyleId { get; set; }
        
        [Required]
        [StringLength(200)]
        public string PatternName { get; set; } = string.Empty;
        
        [Required]
        [StringLength(500)]
        public string Context { get; set; } = string.Empty;
        
        [StringLength(100)]
        public string? FontFamily { get; set; }
        
        public float? FontSize { get; set; }
        public bool? IsBold { get; set; }
        public bool? IsItalic { get; set; }
        public bool? IsUnderline { get; set; }
        
        [StringLength(10)]
        public string? Color { get; set; }
        
        public bool? IsStrikethrough { get; set; }
        public bool? IsAllCaps { get; set; }
        public bool? IsSmallCaps { get; set; }
        
        [StringLength(50)]
        public string? Highlighting { get; set; }
        
        public float? CharacterSpacing { get; set; }
        public bool? HasShadow { get; set; }
        public bool? HasOutline { get; set; }
        
        [StringLength(50)]
        public string? Language { get; set; }
        
        [StringLength(50)]
        public string? Alignment { get; set; }
        
        public float? SpacingBefore { get; set; }
        public float? SpacingAfter { get; set; }
        public float? IndentationLeft { get; set; }
        public float? IndentationRight { get; set; }
        public float? FirstLineIndent { get; set; }
        public float? LineSpacing { get; set; }
        
        [StringLength(1000)]
        public string SampleText { get; set; } = string.Empty;
        
        public int OccurrenceCount { get; set; } = 0;
        
        // Audit fields
        [Required]
        [StringLength(100)]
        public string CreatedBy { get; set; } = string.Empty;
        
        [Required]
        public DateTime CreatedOn { get; set; } = DateTime.UtcNow;
        
        // Navigation property
        public virtual TextStyle TextStyle { get; set; } = null!;
    }
} 