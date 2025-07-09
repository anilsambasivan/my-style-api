using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.DTOs
{
    public class VerificationResultDto
    {
        public int Id { get; set; }
        public int TemplateId { get; set; }
        public string TemplateName { get; set; } = string.Empty;
        public string DocumentName { get; set; } = string.Empty;
        public string DocumentPath { get; set; } = string.Empty;
        public DateTime VerificationDate { get; set; }
        public int TotalMismatches { get; set; }
        public string Status { get; set; } = string.Empty;
        public string ErrorMessage { get; set; } = string.Empty;
        public string CreatedBy { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
        public List<MismatchDto> Mismatches { get; set; } = new List<MismatchDto>();
    }

    public class MismatchDto
    {
        public int Id { get; set; }
        public string ContextKey { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public string StructuralRole { get; set; } = string.Empty;
        public object? Expected { get; set; }
        public object? Actual { get; set; }
        public List<string> MismatchFields { get; set; } = new List<string>();
        public string SampleText { get; set; } = string.Empty;
        public string Severity { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
    }

    public class DocumentVerificationRequest
    {
        [Required]
        public int TemplateId { get; set; }

        [Required]
        public IFormFile Document { get; set; } = null!;

        [Required]
        [StringLength(100)]
        public string CreatedBy { get; set; } = string.Empty;

        public bool StrictMode { get; set; } = true;

        public List<string> IgnoreStyleTypes { get; set; } = new List<string>();
    }

    public class VerificationSummaryDto
    {
        public int TotalVerifications { get; set; }
        public int SuccessfulVerifications { get; set; }
        public int FailedVerifications { get; set; }
        public int TotalMismatches { get; set; }
        public Dictionary<string, int> MismatchesBySeverity { get; set; } = new Dictionary<string, int>();
        public Dictionary<string, int> MismatchesByStyleType { get; set; } = new Dictionary<string, int>();
        public List<VerificationResultDto> RecentVerifications { get; set; } = new List<VerificationResultDto>();
    }

    public class MismatchReportDto
    {
        public string ContextKey { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public string StructuralRole { get; set; } = string.Empty;
        public Dictionary<string, object> Expected { get; set; } = new Dictionary<string, object>();
        public Dictionary<string, object> Actual { get; set; } = new Dictionary<string, object>();
        public List<string> MismatchedFields { get; set; } = new List<string>();
        public string SampleText { get; set; } = string.Empty;
        public string Severity { get; set; } = string.Empty;
        public string RecommendedAction { get; set; } = string.Empty;
    }
} 