using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DocStyleVerify.API.Models
{
    /// <summary>
    /// Represents a default style definition from styles.xml
    /// </summary>
    public class DefaultStyle
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        public int TemplateId { get; set; }
        
        [Required]
        [StringLength(200)]
        public string StyleId { get; set; } = string.Empty;
        
        [Required]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;
        
        [Required]
        [StringLength(50)]
        public string Type { get; set; } = string.Empty; // Paragraph, Character, Table, Numbering
        
        [StringLength(200)]
        public string BasedOn { get; set; } = string.Empty;
        
        [StringLength(200)]
        public string NextStyle { get; set; } = string.Empty;
        
        public bool IsDefault { get; set; } = false;
        
        public bool IsCustom { get; set; } = false;
        
        public int Priority { get; set; } = 0;
        
        public bool IsHidden { get; set; } = false;
        
        public bool IsQuickStyle { get; set; } = false;
        
        // Font properties
        [StringLength(100)]
        public string FontFamily { get; set; } = string.Empty;
        
        public float FontSize { get; set; } = 0;
        
        public bool IsBold { get; set; } = false;
        public bool IsItalic { get; set; } = false;
        public bool IsUnderline { get; set; } = false;
        
        [StringLength(10)]
        public string Color { get; set; } = string.Empty;
        
        // Paragraph properties
        [StringLength(50)]
        public string Alignment { get; set; } = string.Empty;
        
        public float SpacingBefore { get; set; } = 0;
        public float SpacingAfter { get; set; } = 0;
        public float IndentationLeft { get; set; } = 0;
        public float IndentationRight { get; set; } = 0;
        public float FirstLineIndent { get; set; } = 0;
        public float LineSpacing { get; set; } = 0;
        
        // Raw XML for complete style definition
        [Column(TypeName = "text")]
        public string RawXml { get; set; } = string.Empty;
        
        // Audit fields
        [Required]
        [StringLength(100)]
        public string CreatedBy { get; set; } = string.Empty;
        
        [Required]
        public DateTime CreatedOn { get; set; } = DateTime.UtcNow;
        
        [StringLength(100)]
        public string ModifiedBy { get; set; } = string.Empty;
        
        public DateTime? ModifiedOn { get; set; }
        
        // Navigation properties
        [ForeignKey("TemplateId")]
        public virtual Template Template { get; set; } = null!;
    }
} 