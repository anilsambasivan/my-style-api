using Microsoft.AspNetCore.Mvc;
using DocStyleVerify.API.Services;
using DocStyleVerify.API.DTOs;
using DocStyleVerify.API.Data;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class VerificationController : ControllerBase
    {
        private readonly IDocumentProcessingService _documentProcessingService;
        private readonly ITemplateService _templateService;
        private readonly DocStyleVerifyDbContext _context;
        private readonly ILogger<VerificationController> _logger;

        public VerificationController(
            IDocumentProcessingService documentProcessingService,
            ITemplateService templateService,
            DocStyleVerifyDbContext context,
            ILogger<VerificationController> logger)
        {
            _documentProcessingService = documentProcessingService;
            _templateService = templateService;
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Verify a document against a template
        /// </summary>
        /// <param name="request">Document verification request</param>
        /// <returns>Verification result with mismatches</returns>
        [HttpPost("verify")]
        [ProducesResponseType(typeof(VerificationResultDto), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<ActionResult<VerificationResultDto>> VerifyDocument([FromForm] DocumentVerificationRequest request)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            if (request.Document == null || request.Document.Length == 0)
                return BadRequest("Document file is required");

            // Validate file format
            var extension = Path.GetExtension(request.Document.FileName).ToLowerInvariant();
            if (extension != ".docx")
                return BadRequest("Only .docx files are supported");

            try
            {
                // Check if template exists
                if (!await _templateService.TemplateExistsAsync(request.TemplateId))
                    return NotFound($"Template with ID {request.TemplateId} not found");

                // Verify document
                using var documentStream = request.Document.OpenReadStream();
                var result = await _documentProcessingService.VerifyDocumentAsync(
                    documentStream,
                    request.Document.FileName,
                    request.TemplateId,
                    request.CreatedBy,
                    request.StrictMode,
                    request.IgnoreStyleTypes);

                _logger.LogInformation("Document {DocumentName} verified against template {TemplateId}. Found {MismatchCount} mismatches",
                    request.Document.FileName, request.TemplateId, result.TotalMismatches);

                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error verifying document {DocumentName} against template {TemplateId}",
                    request.Document.FileName, request.TemplateId);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get verification results with optional filtering
        /// </summary>
        /// <param name="templateId">Filter by template ID (optional)</param>
        /// <param name="page">Page number (default: 1)</param>
        /// <param name="pageSize">Page size (default: 20)</param>
        /// <param name="status">Filter by status (optional)</param>
        /// <returns>List of verification results</returns>
        [HttpGet("results")]
        [ProducesResponseType(typeof(IEnumerable<VerificationResultDto>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<ActionResult<IEnumerable<VerificationResultDto>>> GetVerificationResults(
            [FromQuery] int? templateId = null,
            [FromQuery] int page = 1,
            [FromQuery] int pageSize = 20,
            [FromQuery] string? status = null)
        {
            if (page < 1)
                return BadRequest("Page must be greater than 0");

            if (pageSize < 1 || pageSize > 100)
                return BadRequest("Page size must be between 1 and 100");

            try
            {
                var query = _context.VerificationResults
                    .Include(vr => vr.Template)
                    .Include(vr => vr.Mismatches)
                    .AsQueryable();

                if (templateId.HasValue)
                    query = query.Where(vr => vr.TemplateId == templateId.Value);

                if (!string.IsNullOrWhiteSpace(status))
                    query = query.Where(vr => vr.Status == status);

                var results = await query
                    .OrderByDescending(vr => vr.VerificationDate)
                    .Skip((page - 1) * pageSize)
                    .Take(pageSize)
                    .ToListAsync();

                var dtos = results.Select(MapVerificationResultToDto);
                return Ok(dtos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving verification results");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get a specific verification result by ID
        /// </summary>
        /// <param name="id">Verification result ID</param>
        /// <returns>Verification result details</returns>
        [HttpGet("results/{id}")]
        [ProducesResponseType(typeof(VerificationResultDto), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<ActionResult<VerificationResultDto>> GetVerificationResult(int id)
        {
            try
            {
                var result = await _context.VerificationResults
                    .Include(vr => vr.Template)
                    .Include(vr => vr.Mismatches)
                    .FirstOrDefaultAsync(vr => vr.Id == id);

                if (result == null)
                    return NotFound($"Verification result with ID {id} not found");

                return Ok(MapVerificationResultToDto(result));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving verification result {VerificationId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get detailed mismatch report for a verification result
        /// </summary>
        /// <param name="id">Verification result ID</param>
        /// <returns>Detailed mismatch report</returns>
        [HttpGet("results/{id}/report")]
        [ProducesResponseType(typeof(IEnumerable<MismatchReportDto>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<ActionResult<IEnumerable<MismatchReportDto>>> GetMismatchReport(int id)
        {
            try
            {
                var result = await _context.VerificationResults
                    .Include(vr => vr.Mismatches)
                    .FirstOrDefaultAsync(vr => vr.Id == id);

                if (result == null)
                    return NotFound($"Verification result with ID {id} not found");

                var reportDtos = result.Mismatches.Select(MapMismatchToReportDto);
                return Ok(reportDtos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating mismatch report for verification {VerificationId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get verification summary statistics
        /// </summary>
        /// <param name="templateId">Filter by template ID (optional)</param>
        /// <param name="fromDate">Filter from date (optional)</param>
        /// <param name="toDate">Filter to date (optional)</param>
        /// <returns>Verification summary</returns>
        [HttpGet("summary")]
        [ProducesResponseType(typeof(VerificationSummaryDto), StatusCodes.Status200OK)]
        public async Task<ActionResult<VerificationSummaryDto>> GetVerificationSummary(
            [FromQuery] int? templateId = null,
            [FromQuery] DateTime? fromDate = null,
            [FromQuery] DateTime? toDate = null)
        {
            try
            {
                var query = _context.VerificationResults
                    .Include(vr => vr.Template)
                    .Include(vr => vr.Mismatches)
                    .AsQueryable();

                if (templateId.HasValue)
                    query = query.Where(vr => vr.TemplateId == templateId.Value);

                if (fromDate.HasValue)
                    query = query.Where(vr => vr.VerificationDate >= fromDate.Value);

                if (toDate.HasValue)
                    query = query.Where(vr => vr.VerificationDate <= toDate.Value);

                var results = await query.ToListAsync();

                var summary = new VerificationSummaryDto
                {
                    TotalVerifications = results.Count,
                    SuccessfulVerifications = results.Count(r => r.Status == "Completed"),
                    FailedVerifications = results.Count(r => r.Status == "Failed"),
                    TotalMismatches = results.Sum(r => r.TotalMismatches),
                    MismatchesBySeverity = results
                        .SelectMany(r => r.Mismatches)
                        .GroupBy(m => m.Severity)
                        .ToDictionary(g => g.Key, g => g.Count()),
                    RecentVerifications = results
                        .OrderByDescending(r => r.VerificationDate)
                        .Take(10)
                        .Select(MapVerificationResultToDto)
                        .ToList()
                };

                return Ok(summary);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating verification summary");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Compare two documents style-wise
        /// </summary>
        /// <param name="request">Document comparison request</param>
        /// <returns>Style comparison results</returns>
        [HttpPost("compare")]
        [ProducesResponseType(typeof(IEnumerable<StyleComparisonDto>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<ActionResult<IEnumerable<StyleComparisonDto>>> CompareDocuments([FromForm] DocumentComparisonRequest request)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            if (request.Document1 == null || request.Document2 == null)
                return BadRequest("Both documents are required");

            try
            {
                // Extract styles from both documents
                using var stream1 = request.Document1.OpenReadStream();
                using var stream2 = request.Document2.OpenReadStream();

                var styles1 = await _documentProcessingService.ExtractDocumentStylesAsync(stream1, request.Document1.FileName);
                var styles2 = await _documentProcessingService.ExtractDocumentStylesAsync(stream2, request.Document2.FileName);

                // Compare styles
                var comparisons = new List<StyleComparisonDto>();
                
                foreach (var style1 in styles1)
                {
                    var matchingStyle2 = styles2.FirstOrDefault(s => s.StyleSignature == style1.StyleSignature);
                    
                    var comparison = new StyleComparisonDto
                    {
                        TemplateStyle = MapTextStyleToDto(style1),
                        DocumentStyle = matchingStyle2 != null ? MapTextStyleToDto(matchingStyle2) : new TextStyleDto(),
                        IsMatch = matchingStyle2 != null,
                        ComparisonSummary = matchingStyle2 != null ? "Styles match" : "Style not found in second document"
                    };

                    if (matchingStyle2 != null)
                    {
                        comparison.DifferentProperties = GetDifferentProperties(style1, matchingStyle2);
                        comparison.IsMatch = comparison.DifferentProperties.Count == 0;
                        comparison.ComparisonSummary = comparison.IsMatch ? "Styles match exactly" : $"Differences in: {string.Join(", ", comparison.DifferentProperties)}";
                    }

                    comparisons.Add(comparison);
                }

                return Ok(comparisons);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error comparing documents {Doc1} and {Doc2}", request.Document1.FileName, request.Document2.FileName);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Delete a verification result
        /// </summary>
        /// <param name="id">Verification result ID</param>
        /// <returns>No content if successful</returns>
        [HttpDelete("results/{id}")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<IActionResult> DeleteVerificationResult(int id)
        {
            try
            {
                var result = await _context.VerificationResults.FindAsync(id);
                if (result == null)
                    return NotFound($"Verification result with ID {id} not found");

                _context.VerificationResults.Remove(result);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Verification result {VerificationId} deleted", id);
                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting verification result {VerificationId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        private VerificationResultDto MapVerificationResultToDto(Models.VerificationResult result)
        {
            return new VerificationResultDto
            {
                Id = result.Id,
                TemplateId = result.TemplateId,
                TemplateName = result.Template?.Name ?? "",
                DocumentName = result.DocumentName,
                DocumentPath = result.DocumentPath,
                VerificationDate = result.VerificationDate,
                TotalMismatches = result.TotalMismatches,
                Status = result.Status,
                ErrorMessage = result.ErrorMessage,
                CreatedBy = result.CreatedBy,
                CreatedOn = result.CreatedOn,
                Mismatches = result.Mismatches?.Select(MapMismatchToDto).ToList() ?? new List<MismatchDto>()
            };
        }

        private MismatchDto MapMismatchToDto(Models.Mismatch mismatch)
        {
            return new MismatchDto
            {
                Id = mismatch.Id,
                ContextKey = mismatch.ContextKey,
                Location = mismatch.Location,
                StructuralRole = mismatch.StructuralRole,
                Expected = !string.IsNullOrEmpty(mismatch.Expected) ? System.Text.Json.JsonSerializer.Deserialize<object>(mismatch.Expected) : null,
                Actual = !string.IsNullOrEmpty(mismatch.Actual) ? System.Text.Json.JsonSerializer.Deserialize<object>(mismatch.Actual) : null,
                MismatchFields = mismatch.MismatchFields.Split(',', StringSplitOptions.RemoveEmptyEntries).ToList(),
                SampleText = mismatch.SampleText,
                Severity = mismatch.Severity,
                CreatedOn = mismatch.CreatedOn
            };
        }

        private MismatchReportDto MapMismatchToReportDto(Models.Mismatch mismatch)
        {
            var expected = !string.IsNullOrEmpty(mismatch.Expected) ? 
                System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(mismatch.Expected) ?? new Dictionary<string, object>() : 
                new Dictionary<string, object>();

            var actual = !string.IsNullOrEmpty(mismatch.Actual) ?
                System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(mismatch.Actual) ?? new Dictionary<string, object>() :
                new Dictionary<string, object>();

            return new MismatchReportDto
            {
                ContextKey = mismatch.ContextKey,
                Location = mismatch.Location,
                StructuralRole = mismatch.StructuralRole,
                Expected = expected,
                Actual = actual,
                MismatchedFields = mismatch.MismatchFields.Split(',', StringSplitOptions.RemoveEmptyEntries).ToList(),
                SampleText = mismatch.SampleText,
                Severity = mismatch.Severity,
                RecommendedAction = GenerateRecommendedAction(mismatch.Severity, mismatch.MismatchFields)
            };
        }

        private TextStyleDto MapTextStyleToDto(Models.TextStyle style)
        {
            return new TextStyleDto
            {
                Id = style.Id,
                TemplateId = style.TemplateId,
                Name = style.Name,
                FontFamily = style.FontFamily,
                FontSize = style.FontSize,
                IsBold = style.IsBold,
                IsItalic = style.IsItalic,
                IsUnderline = style.IsUnderline,
                Color = style.Color,
                Alignment = style.Alignment,
                StyleType = style.StyleType,
                StyleSignature = style.StyleSignature
                // Add other properties as needed
            };
        }

        private List<string> GetDifferentProperties(Models.TextStyle style1, Models.TextStyle style2)
        {
            var differences = new List<string>();

            if (style1.FontFamily != style2.FontFamily) differences.Add("FontFamily");
            if (style1.FontSize != style2.FontSize) differences.Add("FontSize");
            if (style1.IsBold != style2.IsBold) differences.Add("IsBold");
            if (style1.IsItalic != style2.IsItalic) differences.Add("IsItalic");
            if (style1.Color != style2.Color) differences.Add("Color");
            if (style1.Alignment != style2.Alignment) differences.Add("Alignment");
            // Add more property comparisons as needed

            return differences;
        }

        private string GenerateRecommendedAction(string severity, string mismatchFields)
        {
            var fields = mismatchFields.Split(',', StringSplitOptions.RemoveEmptyEntries);
            
            if (severity == "Critical")
                return "Immediate correction required. These styling differences may significantly impact document appearance and compliance.";
            
            if (severity == "High")
                return "Should be corrected. These differences may affect document consistency and professional appearance.";
            
            if (severity == "Medium")
                return "Consider correcting. These differences are noticeable but may not significantly impact document quality.";
            
            return "Optional correction. These are minor styling differences that have minimal impact.";
        }
    }

    public class DocumentComparisonRequest
    {
        [Required]
        public IFormFile Document1 { get; set; } = null!;

        [Required]
        public IFormFile Document2 { get; set; } = null!;

        [Required]
        [StringLength(100)]
        public string CreatedBy { get; set; } = string.Empty;
    }
} 