using Microsoft.AspNetCore.Mvc;
using DocStyleVerify.API.Services;
using DocStyleVerify.API.DTOs;
using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class TemplatesController : ControllerBase
    {
        private readonly ITemplateService _templateService;
        private readonly ILogger<TemplatesController> _logger;

        public TemplatesController(ITemplateService templateService, ILogger<TemplatesController> logger)
        {
            _templateService = templateService;
            _logger = logger;
        }

        /// <summary>
        /// Get all templates with optional filtering and pagination
        /// </summary>
        /// <param name="page">Page number (default: 1)</param>
        /// <param name="pageSize">Page size (default: 20, max: 100)</param>
        /// <param name="status">Filter by status (Active, Inactive, Archived)</param>
        /// <returns>List of templates</returns>
        [HttpGet]
        [ProducesResponseType(typeof(IEnumerable<TemplateDto>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<ActionResult<IEnumerable<TemplateDto>>> GetTemplates(
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
                var templates = await _templateService.GetTemplatesAsync(page, pageSize, status);
                return Ok(templates);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving templates");
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get a specific template by ID
        /// </summary>
        /// <param name="id">Template ID</param>
        /// <returns>Template details</returns>
        [HttpGet("{id}")]
        [ProducesResponseType(typeof(TemplateDto), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<ActionResult<TemplateDto>> GetTemplate(int id)
        {
            try
            {
                var template = await _templateService.GetTemplateAsync(id);
                if (template == null)
                    return NotFound($"Template with ID {id} not found");

                return Ok(template);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving template {TemplateId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Create a new template from uploaded .docx file
        /// </summary>
        /// <param name="createDto">Template creation data</param>
        /// <returns>Created template</returns>
        [HttpPost]
        [ProducesResponseType(typeof(TemplateDto), StatusCodes.Status201Created)]
        [ProducesResponseType(typeof(ValidationProblemDetails), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status409Conflict)]
        public async Task<ActionResult<TemplateDto>> CreateTemplate([FromForm] CreateTemplateDto createDto)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                var template = await _templateService.CreateTemplateAsync(createDto);
                return CreatedAtAction(nameof(GetTemplate), new { id = template.Id }, template);
            }
            catch (ArgumentException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (InvalidOperationException ex)
            {
                return Conflict(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating template {TemplateName}", createDto.Name);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Update an existing template
        /// </summary>
        /// <param name="id">Template ID</param>
        /// <param name="updateDto">Template update data</param>
        /// <returns>Updated template</returns>
        [HttpPut("{id}")]
        [ProducesResponseType(typeof(TemplateDto), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(ValidationProblemDetails), StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status409Conflict)]
        public async Task<ActionResult<TemplateDto>> UpdateTemplate(int id, [FromBody] UpdateTemplateDto updateDto)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            try
            {
                var template = await _templateService.UpdateTemplateAsync(id, updateDto);
                if (template == null)
                    return NotFound($"Template with ID {id} not found");

                return Ok(template);
            }
            catch (InvalidOperationException ex)
            {
                return Conflict(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error updating template {TemplateId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Delete a template
        /// </summary>
        /// <param name="id">Template ID</param>
        /// <returns>No content if successful</returns>
        [HttpDelete("{id}")]
        [ProducesResponseType(StatusCodes.Status204NoContent)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<IActionResult> DeleteTemplate(int id)
        {
            try
            {
                var deleted = await _templateService.DeleteTemplateAsync(id);
                if (!deleted)
                    return NotFound($"Template with ID {id} not found");

                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting template {TemplateId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Process a template to extract all style information
        /// </summary>
        /// <param name="id">Template ID</param>
        /// <returns>Extraction result</returns>
        [HttpPost("{id}/process")]
        [ProducesResponseType(typeof(TemplateExtractionResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<ActionResult<TemplateExtractionResult>> ProcessTemplate(int id)
        {
            try
            {
                var result = await _templateService.ProcessTemplateAsync(id);
                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                return NotFound(ex.Message);
            }
            catch (FileNotFoundException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing template {TemplateId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get all styles for a specific template
        /// </summary>
        /// <param name="id">Template ID</param>
        /// <param name="styleType">Filter by style type (optional)</param>
        /// <returns>List of text styles</returns>
        [HttpGet("{id}/styles")]
        [ProducesResponseType(typeof(IEnumerable<TextStyleDto>), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<ActionResult<IEnumerable<TextStyleDto>>> GetTemplateStyles(
            int id,
            [FromQuery] string? styleType = null)
        {
            try
            {
                if (!await _templateService.TemplateExistsAsync(id))
                    return NotFound($"Template with ID {id} not found");

                var styles = await _templateService.GetTemplateStylesAsync(id, styleType);
                return Ok(styles);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving styles for template {TemplateId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Get template processing status and statistics
        /// </summary>
        /// <param name="id">Template ID</param>
        /// <returns>Template statistics</returns>
        [HttpGet("{id}/stats")]
        [ProducesResponseType(typeof(TemplateStatsDto), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<ActionResult<TemplateStatsDto>> GetTemplateStats(int id)
        {
            try
            {
                var template = await _templateService.GetTemplateAsync(id);
                if (template == null)
                    return NotFound($"Template with ID {id} not found");

                var styles = await _templateService.GetTemplateStylesAsync(id);
                var stylesList = styles.ToList();

                var stats = new TemplateStatsDto
                {
                    TemplateId = id,
                    TemplateName = template.Name,
                    TotalStyles = stylesList.Count,
                    StyleTypeBreakdown = stylesList
                        .GroupBy(s => s.StyleType)
                        .ToDictionary(g => g.Key, g => g.Count()),
                    DirectFormatPatternsCount = stylesList.Sum(s => s.DirectFormatPatterns.Count),
                    LastProcessed = template.ModifiedOn ?? template.CreatedOn,
                    FileSize = template.FileSize,
                    Status = template.Status
                };

                return Ok(stats);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error retrieving stats for template {TemplateId}", id);
                return StatusCode(500, "Internal server error");
            }
        }

        /// <summary>
        /// Validate template file format
        /// </summary>
        /// <param name="file">Template file to validate</param>
        /// <returns>Validation result</returns>
        [HttpPost("validate")]
        [ProducesResponseType(typeof(FileValidationResult), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        public async Task<ActionResult<FileValidationResult>> ValidateTemplateFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("File is required");

            var result = new FileValidationResult
            {
                FileName = file.FileName,
                FileSize = file.Length,
                IsValid = false,
                ValidationMessages = new List<string>()
            };

            try
            {
                // Basic file validation
                var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
                if (extension != ".docx")
                {
                    result.ValidationMessages.Add("Only .docx files are supported");
                    return Ok(result);
                }

                if (file.Length > 50 * 1024 * 1024) // 50MB limit
                {
                    result.ValidationMessages.Add("File size must be less than 50MB");
                    return Ok(result);
                }

                // Document structure validation would go here
                result.IsValid = true;
                result.ValidationMessages.Add("File is valid");

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error validating template file {FileName}", file.FileName);
                result.ValidationMessages.Add("Error validating file");
                return Ok(result);
            }
        }
    }

    public class TemplateStatsDto
    {
        public int TemplateId { get; set; }
        public string TemplateName { get; set; } = string.Empty;
        public int TotalStyles { get; set; }
        public Dictionary<string, int> StyleTypeBreakdown { get; set; } = new Dictionary<string, int>();
        public int DirectFormatPatternsCount { get; set; }
        public DateTime LastProcessed { get; set; }
        public long FileSize { get; set; }
        public string Status { get; set; } = string.Empty;
    }

    public class FileValidationResult
    {
        public string FileName { get; set; } = string.Empty;
        public long FileSize { get; set; }
        public bool IsValid { get; set; }
        public List<string> ValidationMessages { get; set; } = new List<string>();
    }
} 