using DocStyleVerify.API.Models;
using DocStyleVerify.API.DTOs;

namespace DocStyleVerify.API.Services
{
    public interface ITemplateService
    {
        Task<TemplateDto?> GetTemplateAsync(int id);
        Task<IEnumerable<TemplateDto>> GetTemplatesAsync(int page = 1, int pageSize = 20, string? status = null);
        Task<TemplateDto> CreateTemplateAsync(CreateTemplateDto createDto);
        Task<TemplateDto?> UpdateTemplateAsync(int id, UpdateTemplateDto updateDto);
        Task<bool> DeleteTemplateAsync(int id);
        Task<TemplateExtractionResult> ProcessTemplateAsync(int templateId);
        Task<IEnumerable<TextStyleDto>> GetTemplateStylesAsync(int templateId, string? styleType = null);
        Task<TemplateStylesResponseDto> GetTemplateStylesWithDetailsAsync(int templateId, string? styleType = null);
        Task<bool> TemplateExistsAsync(int id);
        Task<string> GetTemplateFilePathAsync(int id);
        Task<Template?> GetTemplateEntityAsync(int id);
    }
} 