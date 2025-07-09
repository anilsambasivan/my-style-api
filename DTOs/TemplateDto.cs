using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.DTOs
{
    public class TemplateDto
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string FileHash { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public string Status { get; set; } = string.Empty;
        public string CreatedBy { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
        public string ModifiedBy { get; set; } = string.Empty;
        public DateTime? ModifiedOn { get; set; }
        public int Version { get; set; }
        public int TextStylesCount { get; set; }
    }

    public class CreateTemplateDto
    {
        [Required]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;

        [StringLength(500)]
        public string Description { get; set; } = string.Empty;

        [Required]
        public IFormFile File { get; set; } = null!;

        [Required]
        [StringLength(100)]
        public string CreatedBy { get; set; } = string.Empty;
    }

    public class UpdateTemplateDto
    {
        [Required]
        [StringLength(200)]
        public string Name { get; set; } = string.Empty;

        [StringLength(500)]
        public string Description { get; set; } = string.Empty;

        [Required]
        [StringLength(20)]
        public string Status { get; set; } = string.Empty;

        [Required]
        [StringLength(100)]
        public string ModifiedBy { get; set; } = string.Empty;
    }

    public class TemplateExtractionResult
    {
        public int TemplateId { get; set; }
        public string TemplateName { get; set; } = string.Empty;
        public int TotalTextStyles { get; set; }
        public int TotalDirectFormatPatterns { get; set; }
        public List<string> ExtractedStyleTypes { get; set; } = new List<string>();
        public List<string> ProcessingMessages { get; set; } = new List<string>();
        public bool Success { get; set; }
        public string ErrorMessage { get; set; } = string.Empty;
        public DateTime ProcessedOn { get; set; } = DateTime.UtcNow;
    }
} 