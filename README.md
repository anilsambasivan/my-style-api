# DocStyleVerify.API

A comprehensive .NET 9.0 Web API for document style verification and template management. This system analyzes Word documents (.docx) to extract styling information and compare them against predefined templates to identify formatting inconsistencies.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Architecture](#architecture)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [API Documentation](#api-documentation)
- [Data Models](#data-models)
- [Usage Examples](#usage-examples)
- [Development](#development)
- [Contributing](#contributing)
- [License](#license)

## Overview

DocStyleVerify.API is designed to ensure document consistency by:
- Extracting comprehensive style information from Word documents
- Creating reusable templates from reference documents
- Verifying documents against templates to identify style mismatches
- Providing detailed reports on formatting inconsistencies
- Supporting enterprise-level document quality assurance workflows

## Features

### Core Functionality
- **Template Management**: Create, update, and manage document templates
- **Style Extraction**: Extract comprehensive styling information from Word documents
- **Document Verification**: Compare documents against templates with detailed mismatch reporting
- **Batch Processing**: Process multiple documents efficiently
- **Flexible Comparison**: Support for strict and lenient comparison modes

### Document Elements Supported
- **Text Formatting**: Font family, size, color, bold, italic, underline, strikethrough
- **Paragraph Formatting**: Alignment, spacing, indentation, line spacing
- **List Formatting**: Numbering, bullets, indentation levels
- **Table Formatting**: Borders, cell spacing, alignment, background colors
- **Page Layout**: Margins, orientation, columns, headers/footers
- **Advanced Features**: Tab stops, direct formatting patterns, hyperlinks, images

### Technical Features
- **PostgreSQL Integration**: Robust data persistence with Entity Framework Core
- **RESTful API**: Clean, well-documented API endpoints
- **Swagger Documentation**: Interactive API documentation
- **File Upload Support**: Secure file handling with validation
- **Audit Trails**: Complete tracking of changes and operations
- **Error Handling**: Comprehensive error handling and logging

## Architecture

### Project Structure
```
DocStyleVerify.API/
├── Controllers/           # API Controllers
│   ├── TemplatesController.cs
│   └── VerificationController.cs
├── Data/                 # Database Context
│   └── DocStyleVerifyDbContext.cs
├── DTOs/                 # Data Transfer Objects
│   ├── StyleDto.cs
│   ├── TemplateDto.cs
│   └── VerificationDto.cs
├── Models/               # Domain Models
│   ├── Template.cs
│   ├── TextStyle.cs
│   ├── VerificationResult.cs
│   ├── DirectFormatPattern.cs
│   ├── FormattingContext.cs
│   └── TabStop.cs
├── Services/             # Business Logic
│   ├── IDocumentProcessingService.cs
│   ├── DocumentProcessingService.cs
│   ├── ITemplateService.cs
│   └── TemplateService.cs
├── Extensions/           # Extension Methods
├── Migrations/           # Entity Framework Migrations
└── Templates/           # Template File Storage
```

### Technology Stack
- **.NET 9.0**: Modern web framework
- **Entity Framework Core 9.0**: ORM with PostgreSQL provider
- **DocumentFormat.OpenXml**: Word document processing
- **Swashbuckle.AspNetCore**: API documentation
- **Serilog**: Structured logging
- **PostgreSQL**: Primary database

## Prerequisites

- .NET 9.0 SDK
- PostgreSQL 12 or higher
- Visual Studio 2022 or VS Code
- Postman or similar API testing tool (optional)

## Installation

### 1. Clone the Repository
```bash
git clone <repository-url>
cd DocStyleVerify.API
```

### 2. Setup Database
```bash
# Install PostgreSQL and create database
createdb DocStyleVerifyDB

# Update connection string in appsettings.json
# Run migrations
dotnet ef database update
```

### 3. Install Dependencies
```bash
dotnet restore
```

### 4. Build and Run
```bash
dotnet build
dotnet run
```

The API will be available at `https://localhost:5001` (or the port specified in launchSettings.json).

## Configuration

### Database Configuration
Update `appsettings.json` with your PostgreSQL connection string:
```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Host=localhost;Database=DocStyleVerifyDB;Username=your_username;Password=your_password"
  }
}
```

### Application Settings
```json
{
  "TemplateStorage": {
    "Path": "Templates"
  },
  "DocumentProcessing": {
    "MaxFileSizeBytes": 104857600,
    "SupportedFormats": ["docx"],
    "ProcessingTimeoutMinutes": 10
  },
  "Verification": {
    "StrictModeByDefault": true,
    "DefaultSeverityMapping": {
      "FontFamily": "Medium",
      "FontSize": "Medium",
      "Color": "High",
      "Alignment": "Medium"
    }
  }
}
```

## API Documentation

### Templates Controller (`/api/templates`)

#### Get All Templates
```http
GET /api/templates?page=1&pageSize=20&status=Active
```
**Response**: List of templates with pagination

#### Get Template by ID
```http
GET /api/templates/{id}
```
**Response**: Template details

#### Create Template
```http
POST /api/templates
Content-Type: multipart/form-data

{
  "name": "Corporate Template",
  "description": "Standard corporate document template", 
  "file": [Word document file],
  "createdBy": "user@example.com"
}
```
**Response**: Created template with ID

#### Update Template
```http
PUT /api/templates/{id}
Content-Type: application/json

{
  "name": "Updated Template Name",
  "description": "Updated description",
  "status": "Active", 
  "modifiedBy": "user@example.com"
}
```

#### Delete Template
```http
DELETE /api/templates/{id}
```

#### Process Template
```http
POST /api/templates/{id}/process
```
**Response**: Extraction result with style information

#### Get Template Styles
```http
GET /api/templates/{id}/styles?styleType=Paragraph
```
**Response**: List of styles for the template

#### Get Template Statistics
```http
GET /api/templates/{id}/stats
```
**Response**: Template statistics and metadata

#### Validate Template File
```http
POST /api/templates/validate
Content-Type: multipart/form-data

file: [Word document file]
```
**Response**: File validation result

### Verification Controller (`/api/verification`)

#### Verify Document
```http
POST /api/verification/verify
Content-Type: multipart/form-data

{
  "document": [Word document file],
  "templateId": 1,
  "createdBy": "user@example.com",
  "strictMode": true,
  "ignoreStyleTypes": ["Character", "List"]
}
```
**Response**: Verification result with mismatches

#### Get Verification Results
```http
GET /api/verification/results?templateId=1&page=1&pageSize=20&status=Completed
```
**Response**: List of verification results

#### Get Verification Result by ID
```http
GET /api/verification/results/{id}
```
**Response**: Detailed verification result

#### Get Mismatch Report
```http
GET /api/verification/results/{id}/report
```
**Response**: Detailed mismatch report

#### Get Verification Summary
```http
GET /api/verification/summary?templateId=1&fromDate=2024-01-01&toDate=2024-12-31
```
**Response**: Verification statistics and summary

#### Compare Documents
```http
POST /api/verification/compare
Content-Type: multipart/form-data

{
  "document1": [Word document file],
  "document2": [Word document file], 
  "createdBy": "user@example.com"
}
```
**Response**: Style comparison result

#### Delete Verification Result
```http
DELETE /api/verification/results/{id}
```

## Data Models

### Template
```csharp
public class Template
{
    public int Id { get; set; }
    public string Name { get; set; }
    public string Description { get; set; }
    public string FileName { get; set; }
    public string FilePath { get; set; }
    public string FileHash { get; set; }
    public long FileSize { get; set; }
    public string Status { get; set; } // Active, Inactive, Archived
    public string CreatedBy { get; set; }
    public DateTime CreatedOn { get; set; }
    public string ModifiedBy { get; set; }
    public DateTime? ModifiedOn { get; set; }
    public int Version { get; set; }
    public virtual ICollection<TextStyle> TextStyles { get; set; }
}
```

### TextStyle
```csharp
public class TextStyle
{
    public int Id { get; set; }
    public int TemplateId { get; set; }
    public string Name { get; set; }
    public string FontFamily { get; set; }
    public float FontSize { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public string Color { get; set; }
    public string Alignment { get; set; }
    public float SpacingBefore { get; set; }
    public float SpacingAfter { get; set; }
    public float IndentationLeft { get; set; }
    public float IndentationRight { get; set; }
    public string StyleType { get; set; } // Paragraph, Character, Table, List, Section
    public string StyleSignature { get; set; }
    // ... additional properties for comprehensive formatting
}
```

### VerificationResult
```csharp
public class VerificationResult
{
    public int Id { get; set; }
    public int TemplateId { get; set; }
    public string DocumentName { get; set; }
    public string DocumentPath { get; set; }
    public DateTime VerificationDate { get; set; }
    public int TotalMismatches { get; set; }
    public string Status { get; set; } // Pending, Completed, Failed
    public string CreatedBy { get; set; }
    public virtual Template Template { get; set; }
    public virtual ICollection<Mismatch> Mismatches { get; set; }
}
```

### Mismatch
```csharp
public class Mismatch
{
    public int Id { get; set; }
    public int VerificationResultId { get; set; }
    public string ContextKey { get; set; }
    public string Location { get; set; }
    public string StructuralRole { get; set; }
    public string Expected { get; set; } // JSON serialized
    public string Actual { get; set; } // JSON serialized
    public string MismatchFields { get; set; }
    public string SampleText { get; set; }
    public string Severity { get; set; } // Low, Medium, High, Critical
}
```

## Usage Examples

### Creating a Template
```bash
curl -X POST "https://localhost:5001/api/templates" \
  -H "Content-Type: multipart/form-data" \
  -F "name=Corporate Standard" \
  -F "description=Standard corporate document template" \
  -F "file=@template.docx" \
  -F "createdBy=admin@company.com"
```

### Verifying a Document
```bash
curl -X POST "https://localhost:5001/api/verification/verify" \
  -H "Content-Type: multipart/form-data" \
  -F "document=@document.docx" \
  -F "templateId=1" \
  -F "createdBy=user@company.com" \
  -F "strictMode=true"
```

### Getting Verification Results
```bash
curl -X GET "https://localhost:5001/api/verification/results?templateId=1&page=1&pageSize=10"
```

## Development

### Running Tests
```bash
dotnet test
```

### Code Style
The project follows standard C# coding conventions and includes:
- XML documentation comments
- Comprehensive error handling
- Logging throughout the application
- Input validation and sanitization

### Database Migrations
```bash
# Add new migration
dotnet ef migrations add MigrationName

# Update database
dotnet ef database update

# Remove last migration
dotnet ef migrations remove
```

### Debugging
The application includes comprehensive logging using Serilog. Log levels can be configured in `appsettings.json`.

## API Features

### Authentication & Authorization
- Currently configured for open access
- Ready for integration with authentication providers
- Audit trail support for all operations

### File Handling
- Secure file upload with validation
- Support for .docx files up to 100MB
- Automatic file cleanup and management

### Error Handling
- Comprehensive exception handling
- Detailed error responses
- Logging of all errors for debugging

### Performance
- Efficient document processing using OpenXML
- Database query optimization
- Pagination support for large datasets

## Troubleshooting

### Common Issues

1. **Database Connection Issues**
   - Verify PostgreSQL is running
   - Check connection string in appsettings.json
   - Ensure database exists and user has permissions

2. **File Upload Issues**
   - Check file size limits in configuration
   - Verify file format is .docx
   - Ensure sufficient disk space

3. **Document Processing Errors**
   - Verify document is not corrupted
   - Check if document is password protected
   - Ensure document uses supported formatting features

### Logging
Application logs are available in the console and can be configured to write to files. Check logs for detailed error information.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## License

[Add your license information here]

---

For more detailed information about specific endpoints, visit the Swagger documentation at `https://localhost:5001` when running the application. 