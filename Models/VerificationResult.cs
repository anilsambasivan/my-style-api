using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.Models
{
    public class VerificationResult
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        public int TemplateId { get; set; }
        
        [Required]
        [StringLength(200)]
        public string DocumentName { get; set; } = string.Empty;
        
        [Required]
        [StringLength(500)]
        public string DocumentPath { get; set; } = string.Empty;
        
        public DateTime VerificationDate { get; set; } = DateTime.UtcNow;
        
        public int TotalMismatches { get; set; } = 0;
        
        [Required]
        [StringLength(20)]
        public string Status { get; set; } = "Completed"; // Pending, Completed, Failed
        
        [StringLength(1000)]
        public string ErrorMessage { get; set; } = string.Empty;
        
        // Audit fields
        [Required]
        [StringLength(100)]
        public string CreatedBy { get; set; } = string.Empty;
        
        [Required]
        public DateTime CreatedOn { get; set; } = DateTime.UtcNow;
        
        // Navigation properties
        public virtual Template Template { get; set; } = null!;
        public virtual ICollection<Mismatch> Mismatches { get; set; } = new List<Mismatch>();
    }

    public class Mismatch
    {
        [Key]
        public int Id { get; set; }
        
        [Required]
        public int VerificationResultId { get; set; }
        
        [Required]
        [StringLength(500)]
        public string ContextKey { get; set; } = string.Empty;
        
        [Required]
        [StringLength(1000)]
        public string Location { get; set; } = string.Empty;
        
        [Required]
        [StringLength(100)]
        public string StructuralRole { get; set; } = string.Empty;
        
        [StringLength(2000)]
        public string Expected { get; set; } = string.Empty; // JSON serialized
        
        [StringLength(2000)]
        public string Actual { get; set; } = string.Empty; // JSON serialized
        
        [Required]
        [StringLength(1000)]
        public string MismatchFields { get; set; } = string.Empty; // Comma separated list
        
        [StringLength(1000)]
        public string SampleText { get; set; } = string.Empty;
        
        [Required]
        [StringLength(50)]
        public string Severity { get; set; } = "Medium"; // Low, Medium, High, Critical
        
        // Audit fields
        [Required]
        public DateTime CreatedOn { get; set; } = DateTime.UtcNow;
        
        // Navigation property
        public virtual VerificationResult VerificationResult { get; set; } = null!;
    }
} 