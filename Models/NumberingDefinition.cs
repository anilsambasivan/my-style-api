using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DocStyleVerify.API.Models
{
    /// <summary>
    /// Represents a numbering definition from numbering.xml
    /// </summary>
    public class NumberingDefinition
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        public int TemplateId { get; set; }
        
        [Required]
        public int AbstractNumId { get; set; }
        
        [Required]
        public int NumberingId { get; set; }
        
        [Required]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;
        
        [StringLength(50)]
        public string Type { get; set; } = string.Empty; // Bullet, Number, Outline, etc.
        
        // Raw XML for complete numbering definition
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
        
        public virtual ICollection<NumberingLevel> NumberingLevels { get; set; } = new List<NumberingLevel>();
    }
    
    /// <summary>
    /// Represents a level within a numbering definition
    /// </summary>
    public class NumberingLevel
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        public int NumberingDefinitionId { get; set; }
        
        [Required]
        public int Level { get; set; }
        
        [StringLength(50)]
        public string NumberFormat { get; set; } = string.Empty; // decimal, bullet, lowerLetter, etc.
        
        [StringLength(200)]
        public string LevelText { get; set; } = string.Empty;
        
        [StringLength(10)]
        public string LevelJustification { get; set; } = string.Empty; // left, center, right
        
        public int StartValue { get; set; } = 1;
        
        public bool IsLegal { get; set; } = false;
        
        // Font properties for numbering
        [StringLength(100)]
        public string FontFamily { get; set; } = string.Empty;
        
        public float FontSize { get; set; } = 0;
        
        public bool IsBold { get; set; } = false;
        public bool IsItalic { get; set; } = false;
        
        [StringLength(10)]
        public string Color { get; set; } = string.Empty;
        
        // Paragraph properties for numbering
        public float IndentationLeft { get; set; } = 0;
        public float IndentationHanging { get; set; } = 0;
        public float TabStopPosition { get; set; } = 0;
        
        // Raw XML for complete level definition
        [Column(TypeName = "text")]
        public string RawXml { get; set; } = string.Empty;
        
        // Navigation properties
        [ForeignKey("NumberingDefinitionId")]
        public virtual NumberingDefinition NumberingDefinition { get; set; } = null!;
    }
} 