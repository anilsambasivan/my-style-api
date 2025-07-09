using DocStyleVerify.API.Models;
using DocStyleVerify.API.DTOs;

namespace DocStyleVerify.API.Services
{
    public interface IDocumentProcessingService
    {
        /// <summary>
        /// Extracts all style information from a Word document template
        /// </summary>
        Task<TemplateExtractionResult> ExtractTemplateStylesAsync(Stream documentStream, string fileName, int templateId, string createdBy);

        /// <summary>
        /// Extracts styles from a document without saving to database. Used for both template extraction and verification.
        /// </summary>
        /// <param name="documentStream">Document stream to extract from</param>
        /// <param name="templateId">Template ID (use 0 for verification scenarios)</param>
        /// <param name="createdBy">User who initiated the extraction</param>
        /// <returns>Tuple containing extracted styles and patterns</returns>
        Task<(List<TextStyle> styles, List<DirectFormatPattern> patterns)> ExtractStylesOnlyAsync(Stream documentStream, int templateId, string createdBy);

        /// <summary>
        /// Verifies a document against a template and returns mismatches
        /// </summary>
        Task<VerificationResultDto> VerifyDocumentAsync(Stream documentStream, string fileName, int templateId, string createdBy, bool strictMode = true, List<string>? ignoreStyleTypes = null);

        /// <summary>
        /// Extracts effective styles from document elements
        /// </summary>
        Task<List<TextStyle>> ExtractDocumentStylesAsync(Stream documentStream, string fileName);

        /// <summary>
        /// Compares two style sets and identifies mismatches
        /// </summary>
        List<MismatchReportDto> CompareStyles(List<TextStyle> templateStyles, List<TextStyle> documentStyles, bool strictMode = true);

        /// <summary>
        /// Generates a style signature for fast comparison
        /// </summary>
        string GenerateStyleSignature(TextStyle style);

        /// <summary>
        /// Validates document format and structure
        /// </summary>
        Task<bool> ValidateDocumentFormatAsync(Stream documentStream);

        /// <summary>
        /// Gets document metadata and statistics
        /// </summary>
        Task<DocumentMetadata> GetDocumentMetadataAsync(Stream documentStream);
    }

    public class DocumentMetadata
    {
        public string Title { get; set; } = string.Empty;
        public string Author { get; set; } = string.Empty;
        public string Subject { get; set; } = string.Empty;
        public DateTime? Created { get; set; }
        public DateTime? LastModified { get; set; }
        public int PageCount { get; set; }
        public int WordCount { get; set; }
        public int ParagraphCount { get; set; }
        public int TableCount { get; set; }
        public int ImageCount { get; set; }
        public List<string> StyleNames { get; set; } = new List<string>();
        public Dictionary<string, int> StyleUsageCounts { get; set; } = new Dictionary<string, int>();
    }
} 