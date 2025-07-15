using DocStyleVerify.API.Models;
using DocStyleVerify.API.DTOs;
using DocStyleVerify.API.Data;
using Microsoft.EntityFrameworkCore;
using System.Security.Cryptography;

namespace DocStyleVerify.API.Services
{
    public class TemplateService : ITemplateService
    {
        private readonly DocStyleVerifyDbContext _context;
        private readonly IDocumentProcessingService _documentProcessingService;
        private readonly IConfiguration _configuration;
        private readonly ILogger<TemplateService> _logger;
        private readonly string _templateStoragePath;

        public TemplateService(
            DocStyleVerifyDbContext context, 
            IDocumentProcessingService documentProcessingService,
            IConfiguration configuration,
            ILogger<TemplateService> logger)
        {
            _context = context;
            _documentProcessingService = documentProcessingService;
            _configuration = configuration;
            _logger = logger;
            _templateStoragePath = _configuration["TemplateStorage:Path"] ?? Path.Combine(Directory.GetCurrentDirectory(), "Templates");
            
            // Ensure template directory exists
            Directory.CreateDirectory(_templateStoragePath);
        }

        public async Task<TemplateDto?> GetTemplateAsync(int id)
        {
            var template = await _context.Templates
                .Include(t => t.TextStyles)
                .FirstOrDefaultAsync(t => t.Id == id);

            if (template == null)
                return null;

            return MapToDto(template);
        }

        public async Task<IEnumerable<TemplateDto>> GetTemplatesAsync(int page = 1, int pageSize = 20, string? status = null)
        {
            var query = _context.Templates.Include(t => t.TextStyles).AsQueryable();

            if (!string.IsNullOrWhiteSpace(status))
            {
                query = query.Where(t => t.Status == status);
            }

            var templates = await query
                .OrderByDescending(t => t.CreatedOn)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            return templates.Select(MapToDto);
        }

        public async Task<TemplateDto> CreateTemplateAsync(CreateTemplateDto createDto)
        {
            // Validate file
            if (createDto.File == null || createDto.File.Length == 0)
                throw new ArgumentException("Template file is required");

            if (!IsValidDocxFile(createDto.File))
                throw new ArgumentException("Only .docx files are supported");

            // Check for duplicate name
            if (await _context.Templates.AnyAsync(t => t.Name == createDto.Name))
                throw new InvalidOperationException($"Template with name '{createDto.Name}' already exists");

            // Calculate file hash
            string fileHash;
            using (var stream = createDto.File.OpenReadStream())
            {
                fileHash = await CalculateFileHashAsync(stream);
            }

            // Check for duplicate file
            if (await _context.Templates.AnyAsync(t => t.FileHash == fileHash))
                throw new InvalidOperationException("A template with identical content already exists");

            // Save file to storage
            var fileName = $"{Guid.NewGuid()}{Path.GetExtension(createDto.File.FileName)}";
            var filePath = Path.Combine(_templateStoragePath, fileName);

            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                await createDto.File.CopyToAsync(fileStream);
            }

            // Create template entity
            var template = new Template
            {
                Name = createDto.Name,
                Description = createDto.Description,
                FileName = createDto.File.FileName,
                FilePath = filePath,
                FileHash = fileHash,
                FileSize = createDto.File.Length,
                Status = "Active",
                CreatedBy = createDto.CreatedBy,
                CreatedOn = DateTime.UtcNow,
                Version = 1
            };

            _context.Templates.Add(template);
            await _context.SaveChangesAsync();

            _logger.LogInformation("Template {TemplateName} created with ID {TemplateId}", template.Name, template.Id);

            return MapToDto(template);
        }

        public async Task<TemplateDto?> UpdateTemplateAsync(int id, UpdateTemplateDto updateDto)
        {
            var template = await _context.Templates.FindAsync(id);
            if (template == null)
                return null;

            // Check for duplicate name (excluding current template)
            if (await _context.Templates.AnyAsync(t => t.Name == updateDto.Name && t.Id != id))
                throw new InvalidOperationException($"Template with name '{updateDto.Name}' already exists");

            template.Name = updateDto.Name;
            template.Description = updateDto.Description;
            template.Status = updateDto.Status;
            template.ModifiedBy = updateDto.ModifiedBy;
            template.ModifiedOn = DateTime.UtcNow;
            template.Version++;

            await _context.SaveChangesAsync();

            _logger.LogInformation("Template {TemplateId} updated by {ModifiedBy}", id, updateDto.ModifiedBy);

            return MapToDto(template);
        }

        public async Task<bool> DeleteTemplateAsync(int id)
        {
            var template = await _context.Templates
                .Include(t => t.TextStyles)
                .FirstOrDefaultAsync(t => t.Id == id);

            if (template == null)
                return false;

            // Delete associated file
            try
            {
                if (File.Exists(template.FilePath))
                {
                    File.Delete(template.FilePath);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Failed to delete template file {FilePath}", template.FilePath);
            }

            // Remove from database (cascade delete will handle related records)
            _context.Templates.Remove(template);
            await _context.SaveChangesAsync();

            _logger.LogInformation("Template {TemplateId} deleted", id);

            return true;
        }

        public async Task<TemplateExtractionResult> ProcessTemplateAsync(int templateId)
        {
            var template = await _context.Templates.FindAsync(templateId);
            if (template == null)
                throw new ArgumentException($"Template with ID {templateId} not found");

            if (!File.Exists(template.FilePath))
                throw new FileNotFoundException($"Template file not found: {template.FilePath}");

            try
            {
                using var fileStream = new FileStream(template.FilePath, FileMode.Open, FileAccess.Read);
                var result = await _documentProcessingService.ExtractTemplateStylesAsync(
                    fileStream, 
                    template.FileName, 
                    templateId, 
                    "System");

                _logger.LogInformation("Template {TemplateId} processed successfully. Extracted {StyleCount} styles", 
                    templateId, result.TotalTextStyles);

                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to process template {TemplateId}", templateId);
                throw;
            }
        }

        public async Task<IEnumerable<TextStyleDto>> GetTemplateStylesAsync(int templateId, string? styleType = null)
        {
            var query = _context.TextStyles
                .Include(ts => ts.FormattingContext)
                .Include(ts => ts.DirectFormatPatterns)
                .Include(ts => ts.TabStops)
                .Where(ts => ts.TemplateId == templateId);

            if (!string.IsNullOrWhiteSpace(styleType))
            {
                query = query.Where(ts => ts.StyleType == styleType);
            }

            var styles = await query.ToListAsync();

            return styles.Select(MapTextStyleToDto);
        }

        public async Task<TemplateStylesResponseDto> GetTemplateStylesWithDetailsAsync(int templateId, string? styleType = null)
        {
            // Get regular styles
            var styles = await GetTemplateStylesAsync(templateId, styleType);

            // Get default styles
            var defaultStylesQuery = _context.DefaultStyles
                .Where(ds => ds.TemplateId == templateId);
            
            if (!string.IsNullOrWhiteSpace(styleType))
            {
                defaultStylesQuery = defaultStylesQuery.Where(ds => ds.Type == styleType);
            }

            var defaultStyles = await defaultStylesQuery.ToListAsync();

            // Get numbering definitions
            var numberingDefinitions = await _context.NumberingDefinitions
                .Include(nd => nd.NumberingLevels)
                .Where(nd => nd.TemplateId == templateId)
                .ToListAsync();

            // Map to DTOs
            var defaultStyleDtos = defaultStyles.Select(MapDefaultStyleToDto);
            var numberingDefinitionDtos = numberingDefinitions.Select(MapNumberingDefinitionToDto);

            return new TemplateStylesResponseDto
            {
                Styles = styles.ToList(),
                DefaultStyles = defaultStyleDtos.ToList(),
                NumberingDefinitions = numberingDefinitionDtos.ToList(),
                TotalStyles = styles.Count(),
                TotalDefaultStyles = defaultStyles.Count,
                TotalNumberingDefinitions = numberingDefinitions.Count,
                ExtractedOn = DateTime.UtcNow
            };
        }

        public async Task<bool> TemplateExistsAsync(int id)
        {
            return await _context.Templates.AnyAsync(t => t.Id == id);
        }

        public async Task<string> GetTemplateFilePathAsync(int id)
        {
            var template = await _context.Templates.FindAsync(id);
            return template?.FilePath ?? string.Empty;
        }

        public async Task<Template?> GetTemplateEntityAsync(int id)
        {
            return await _context.Templates
                .Include(t => t.TextStyles)
                .ThenInclude(ts => ts.FormattingContext)
                .Include(t => t.TextStyles)
                .ThenInclude(ts => ts.DirectFormatPatterns)
                .Include(t => t.TextStyles)
                .ThenInclude(ts => ts.TabStops)
                .FirstOrDefaultAsync(t => t.Id == id);
        }

        private TemplateDto MapToDto(Template template)
        {
            return new TemplateDto
            {
                Id = template.Id,
                Name = template.Name,
                Description = template.Description,
                FileName = template.FileName,
                FilePath = template.FilePath,
                FileHash = template.FileHash,
                FileSize = template.FileSize,
                Status = template.Status,
                CreatedBy = template.CreatedBy,
                CreatedOn = template.CreatedOn,
                ModifiedBy = template.ModifiedBy,
                ModifiedOn = template.ModifiedOn,
                Version = template.Version,
                TextStylesCount = template.TextStyles?.Count ?? 0
            };
        }

        private TextStyleDto MapTextStyleToDto(TextStyle style)
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
                IsStrikethrough = style.IsStrikethrough,
                IsAllCaps = style.IsAllCaps,
                IsSmallCaps = style.IsSmallCaps,
                HasShadow = style.HasShadow,
                HasOutline = style.HasOutline,
                Highlighting = style.Highlighting,
                CharacterSpacing = style.CharacterSpacing,
                Language = style.Language,
                Color = style.Color,
                Alignment = style.Alignment,
                SpacingBefore = style.SpacingBefore,
                SpacingAfter = style.SpacingAfter,
                IndentationLeft = style.IndentationLeft,
                IndentationRight = style.IndentationRight,
                FirstLineIndent = style.FirstLineIndent,
                WidowOrphanControl = style.WidowOrphanControl,
                KeepWithNext = style.KeepWithNext,
                KeepTogether = style.KeepTogether,
                BorderStyle = style.BorderStyle,
                BorderColor = style.BorderColor,
                BorderWidth = style.BorderWidth,
                BorderDirections = style.BorderDirections,
                ListNumberStyle = style.ListNumberStyle,
                ListLevel = style.ListLevel,
                LineSpacing = style.LineSpacing,
                StyleType = style.StyleType,
                BasedOnStyle = style.BasedOnStyle,
                Priority = style.Priority,
                StyleSignature = style.StyleSignature,
                CreatedBy = style.CreatedBy,
                CreatedOn = style.CreatedOn,
                ModifiedBy = style.ModifiedBy,
                ModifiedOn = style.ModifiedOn,
                Version = style.Version,
                FormattingContext = MapFormattingContextToDto(style.FormattingContext),
                DirectFormatPatterns = style.DirectFormatPatterns?.Select(MapDirectFormatPatternToDto).ToList() ?? new List<DirectFormatPatternDto>(),
                TabStops = style.TabStops?.Select(MapTabStopToDto).ToList() ?? new List<TabStopDto>()
            };
        }

        private FormattingContextDto MapFormattingContextToDto(FormattingContext context)
        {
            return new FormattingContextDto
            {
                Id = context.Id,
                StyleName = context.StyleName,
                ElementType = context.ElementType,
                ParentContext = context.ParentContext,
                StructuralRole = context.StructuralRole,
                SampleText = context.SampleText,
                ContentControlTag = context.ContentControlTag,
                ContentControlType = context.ContentControlType,
                SectionIndex = context.SectionIndex,
                TableIndex = context.TableIndex,
                RowIndex = context.RowIndex,
                CellIndex = context.CellIndex,
                ParagraphIndex = context.ParagraphIndex,
                RunIndex = context.RunIndex,
                ContextKey = context.ContextKey,
                ContentControlProperties = context.ContentControlProperties
            };
        }

        private DirectFormatPatternDto MapDirectFormatPatternToDto(DirectFormatPattern pattern)
        {
            return new DirectFormatPatternDto
            {
                Id = pattern.Id,
                TextStyleId = pattern.TextStyleId,
                PatternName = pattern.PatternName,
                Context = pattern.Context,
                FontFamily = pattern.FontFamily,
                FontSize = pattern.FontSize,
                IsBold = pattern.IsBold,
                IsItalic = pattern.IsItalic,
                IsUnderline = pattern.IsUnderline,
                Color = pattern.Color,
                SampleText = pattern.SampleText,
                OccurrenceCount = pattern.OccurrenceCount,
                CreatedBy = pattern.CreatedBy,
                CreatedOn = pattern.CreatedOn
            };
        }

        private TabStopDto MapTabStopToDto(TabStop tabStop)
        {
            return new TabStopDto
            {
                Id = tabStop.Id,
                Position = tabStop.Position,
                Alignment = tabStop.Alignment,
                Leader = tabStop.Leader
            };
        }

        private bool IsValidDocxFile(IFormFile file)
        {
            if (file == null) return false;
            
            var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
            return extension == ".docx";
        }

        private async Task<string> CalculateFileHashAsync(Stream stream)
        {
            using var sha256 = SHA256.Create();
            var hashBytes = await sha256.ComputeHashAsync(stream);
            return Convert.ToBase64String(hashBytes);
        }

        private DefaultStyleDto MapDefaultStyleToDto(DefaultStyle defaultStyle)
        {
            return new DefaultStyleDto
            {
                Id = defaultStyle.Id,
                TemplateId = defaultStyle.TemplateId,
                StyleId = defaultStyle.StyleId,
                Name = defaultStyle.Name,
                Type = defaultStyle.Type,
                BasedOn = defaultStyle.BasedOn,
                NextStyle = defaultStyle.NextStyle,
                IsDefault = defaultStyle.IsDefault,
                IsCustom = defaultStyle.IsCustom,
                Priority = defaultStyle.Priority,
                IsHidden = defaultStyle.IsHidden,
                IsQuickStyle = defaultStyle.IsQuickStyle,
                FontFamily = defaultStyle.FontFamily,
                FontSize = defaultStyle.FontSize,
                IsBold = defaultStyle.IsBold,
                IsItalic = defaultStyle.IsItalic,
                IsUnderline = defaultStyle.IsUnderline,
                Color = defaultStyle.Color,
                Alignment = defaultStyle.Alignment,
                SpacingBefore = defaultStyle.SpacingBefore,
                SpacingAfter = defaultStyle.SpacingAfter,
                IndentationLeft = defaultStyle.IndentationLeft,
                IndentationRight = defaultStyle.IndentationRight,
                FirstLineIndent = defaultStyle.FirstLineIndent,
                LineSpacing = defaultStyle.LineSpacing,
                RawXml = defaultStyle.RawXml,
                CreatedBy = defaultStyle.CreatedBy,
                CreatedOn = defaultStyle.CreatedOn,
                ModifiedBy = defaultStyle.ModifiedBy,
                ModifiedOn = defaultStyle.ModifiedOn
            };
        }

        private NumberingDefinitionDto MapNumberingDefinitionToDto(NumberingDefinition numberingDefinition)
        {
            return new NumberingDefinitionDto
            {
                Id = numberingDefinition.Id,
                TemplateId = numberingDefinition.TemplateId,
                AbstractNumId = numberingDefinition.AbstractNumId,
                NumberingId = numberingDefinition.NumberingId,
                Name = numberingDefinition.Name,
                Type = numberingDefinition.Type,
                RawXml = numberingDefinition.RawXml,
                NumberingLevels = numberingDefinition.NumberingLevels.Select(MapNumberingLevelToDto).ToList(),
                CreatedBy = numberingDefinition.CreatedBy,
                CreatedOn = numberingDefinition.CreatedOn,
                ModifiedBy = numberingDefinition.ModifiedBy,
                ModifiedOn = numberingDefinition.ModifiedOn
            };
        }

        private NumberingLevelDto MapNumberingLevelToDto(NumberingLevel numberingLevel)
        {
            return new NumberingLevelDto
            {
                Id = numberingLevel.Id,
                NumberingDefinitionId = numberingLevel.NumberingDefinitionId,
                Level = numberingLevel.Level,
                NumberFormat = numberingLevel.NumberFormat,
                LevelText = numberingLevel.LevelText,
                LevelJustification = numberingLevel.LevelJustification,
                StartValue = numberingLevel.StartValue,
                IsLegal = numberingLevel.IsLegal,
                FontFamily = numberingLevel.FontFamily,
                FontSize = numberingLevel.FontSize,
                IsBold = numberingLevel.IsBold,
                IsItalic = numberingLevel.IsItalic,
                Color = numberingLevel.Color,
                IndentationLeft = numberingLevel.IndentationLeft,
                IndentationHanging = numberingLevel.IndentationHanging,
                TabStopPosition = numberingLevel.TabStopPosition,
                RawXml = numberingLevel.RawXml
            };
        }
    }
} 