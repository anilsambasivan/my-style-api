using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.Models
{
    public class Template
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;
        
        [StringLength(500)]
        public string Description { get; set; } = string.Empty;
        
        [Required]
        [StringLength(200)]
        public string FileName { get; set; } = string.Empty;
        
        [Required]
        [StringLength(500)]
        public string FilePath { get; set; } = string.Empty;
        
        [Required]
        [StringLength(50)]
        public string FileHash { get; set; } = string.Empty;
        
        public long FileSize { get; set; }
        
        [Required]
        [StringLength(20)]
        public string Status { get; set; } = "Active"; // Active, Inactive, Archived
        
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
        public virtual ICollection<TextStyle> TextStyles { get; set; } = new List<TextStyle>();
        public virtual ICollection<DefaultStyle> DefaultStyles { get; set; } = new List<DefaultStyle>();
        public virtual ICollection<NumberingDefinition> NumberingDefinitions { get; set; } = new List<NumberingDefinition>();
    }
} 