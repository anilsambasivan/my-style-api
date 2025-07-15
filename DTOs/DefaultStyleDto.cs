using System.ComponentModel.DataAnnotations;

namespace DocStyleVerify.API.DTOs
{
    /// <summary>
    /// DTO for default style information from styles.xml
    /// </summary>
    public class DefaultStyleDto
    {
        public int Id { get; set; }
        public int TemplateId { get; set; }
        public string StyleId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string BasedOn { get; set; } = string.Empty;
        public string NextStyle { get; set; } = string.Empty;
        public bool IsDefault { get; set; }
        public bool IsCustom { get; set; }
        public int Priority { get; set; }
        public bool IsHidden { get; set; }
        public bool IsQuickStyle { get; set; }
        
        // Font properties
        public string FontFamily { get; set; } = string.Empty;
        public float FontSize { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public string Color { get; set; } = string.Empty;
        
        // Paragraph properties
        public string Alignment { get; set; } = string.Empty;
        public float SpacingBefore { get; set; }
        public float SpacingAfter { get; set; }
        public float IndentationLeft { get; set; }
        public float IndentationRight { get; set; }
        public float FirstLineIndent { get; set; }
        public float LineSpacing { get; set; }
        
        public string RawXml { get; set; } = string.Empty;
        public string CreatedBy { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
        public string ModifiedBy { get; set; } = string.Empty;
        public DateTime? ModifiedOn { get; set; }
    }
    
    /// <summary>
    /// DTO for numbering definition information from numbering.xml
    /// </summary>
    public class NumberingDefinitionDto
    {
        public int Id { get; set; }
        public int TemplateId { get; set; }
        public int AbstractNumId { get; set; }
        public int NumberingId { get; set; }
        public string Name { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string RawXml { get; set; } = string.Empty;
        public List<NumberingLevelDto> NumberingLevels { get; set; } = new List<NumberingLevelDto>();
        public string CreatedBy { get; set; } = string.Empty;
        public DateTime CreatedOn { get; set; }
        public string ModifiedBy { get; set; } = string.Empty;
        public DateTime? ModifiedOn { get; set; }
    }
    
    /// <summary>
    /// DTO for numbering level information
    /// </summary>
    public class NumberingLevelDto
    {
        public int Id { get; set; }
        public int NumberingDefinitionId { get; set; }
        public int Level { get; set; }
        public string NumberFormat { get; set; } = string.Empty;
        public string LevelText { get; set; } = string.Empty;
        public string LevelJustification { get; set; } = string.Empty;
        public int StartValue { get; set; }
        public bool IsLegal { get; set; }
        
        // Font properties
        public string FontFamily { get; set; } = string.Empty;
        public float FontSize { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public string Color { get; set; } = string.Empty;
        
        // Paragraph properties
        public float IndentationLeft { get; set; }
        public float IndentationHanging { get; set; }
        public float TabStopPosition { get; set; }
        
        public string RawXml { get; set; } = string.Empty;
    }
    
    /// <summary>
    /// Enhanced response DTO for GetTemplateStyles API that includes all style information
    /// </summary>
    public class TemplateStylesResponseDto
    {
        public List<TextStyleDto> Styles { get; set; } = new List<TextStyleDto>();
        public List<DefaultStyleDto> DefaultStyles { get; set; } = new List<DefaultStyleDto>();
        public List<NumberingDefinitionDto> NumberingDefinitions { get; set; } = new List<NumberingDefinitionDto>();
        public int TotalStyles { get; set; }
        public int TotalDefaultStyles { get; set; }
        public int TotalNumberingDefinitions { get; set; }
        public DateTime ExtractedOn { get; set; }
    }
} 