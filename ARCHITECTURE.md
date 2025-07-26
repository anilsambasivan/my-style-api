# Technical Architecture

This document provides a detailed technical overview of the DocStyleVerify.API system architecture, design patterns, and implementation details.

## Table of Contents

- [System Overview](#system-overview)
- [Architecture Patterns](#architecture-patterns)
- [Component Architecture](#component-architecture)
- [Data Architecture](#data-architecture)
- [Processing Pipeline](#processing-pipeline)
- [API Design](#api-design)
- [Security Architecture](#security-architecture)
- [Performance Considerations](#performance-considerations)
- [Scalability Design](#scalability-design)
- [Integration Points](#integration-points)

## System Overview

DocStyleVerify.API is a document style verification system built using a layered architecture pattern with clear separation of concerns. The system processes Microsoft Word documents (.docx) to extract comprehensive styling information and compare them against predefined templates.

### High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Presentation Layer                        │
├─────────────────────────────────────────────────────────────┤
│ Controllers │ DTOs │ Swagger │ Error Handling │ Validation  │
├─────────────────────────────────────────────────────────────┤
│                    Business Logic Layer                      │
├─────────────────────────────────────────────────────────────┤
│ Services │ Document Processing │ Style Extraction │ Comparison│
├─────────────────────────────────────────────────────────────┤
│                    Data Access Layer                         │
├─────────────────────────────────────────────────────────────┤
│ Entity Framework │ Repository Pattern │ Database Context    │
├─────────────────────────────────────────────────────────────┤
│                    Infrastructure Layer                      │
├─────────────────────────────────────────────────────────────┤
│ PostgreSQL │ File System │ Logging │ Configuration │ OpenXML │
└─────────────────────────────────────────────────────────────┘
```

### Key Architectural Principles

1. **Separation of Concerns**: Clear boundaries between presentation, business logic, and data access
2. **Dependency Injection**: Loose coupling through DI container
3. **Single Responsibility**: Each component has a focused responsibility
4. **Open/Closed Principle**: Extensible through interfaces without modifying existing code
5. **Database-First Approach**: Schema-driven design with Entity Framework migrations

## Architecture Patterns

### 1. Layered Architecture

The system follows a traditional layered architecture with four distinct layers:

#### Presentation Layer
- **Controllers**: Handle HTTP requests and responses
- **DTOs**: Data transfer objects for API communication
- **Validators**: Input validation and sanitization
- **Swagger**: API documentation and testing interface

#### Business Logic Layer
- **Services**: Core business logic implementation
- **Processors**: Document and style processing algorithms
- **Comparators**: Style comparison and mismatch detection
- **Validators**: Business rule validation

#### Data Access Layer
- **DbContext**: Entity Framework database context
- **Models**: Domain entities representing business objects
- **Repositories**: Data access abstraction (implicit through EF)
- **Migrations**: Database schema versioning

#### Infrastructure Layer
- **Database**: PostgreSQL for data persistence
- **File System**: Template and document storage
- **Logging**: Structured logging with Serilog
- **Configuration**: Environment-specific settings

### 2. Service-Based Architecture

Services are organized around business capabilities:

```
IDocumentProcessingService
├── ExtractTemplateStylesAsync()
├── ExtractStylesOnlyAsync()
├── VerifyDocumentAsync()
├── ExtractDocumentStylesAsync()
├── CompareStyles()
├── GenerateStyleSignature()
├── ValidateDocumentFormatAsync()
└── GetDocumentMetadataAsync()

ITemplateService
├── GetTemplatesAsync()
├── GetTemplateAsync()
├── CreateTemplateAsync()
├── UpdateTemplateAsync()
├── DeleteTemplateAsync()
├── ProcessTemplateAsync()
└── TemplateExistsAsync()
```

### 3. Repository Pattern (Implicit)

Entity Framework Core acts as a generic repository with:
- **DbSet<T>**: Generic repository for each entity
- **ChangeTracking**: Automatic entity state management
- **Query Translation**: LINQ to SQL conversion
- **Unit of Work**: DbContext as transaction boundary

## Component Architecture

### 1. Document Processing Pipeline

```
Document Upload
      ↓
File Validation
      ↓
OpenXML Processing
      ↓
Style Extraction
      ↓
Style Normalization
      ↓
Database Storage
      ↓
Response Generation
```

### 2. Style Extraction Components

#### OpenXML Parser
- **WordprocessingDocument**: Document container
- **MainDocumentPart**: Document content
- **StyleDefinitionsPart**: Style definitions
- **NumberingDefinitionsPart**: List styles
- **ThemePart**: Document theme

#### Style Processors
- **ParagraphStyleProcessor**: Paragraph formatting
- **CharacterStyleProcessor**: Character formatting
- **TableStyleProcessor**: Table formatting
- **ListStyleProcessor**: List formatting
- **SectionStyleProcessor**: Page/section formatting

#### Context Analyzers
- **StructuralAnalyzer**: Document structure
- **ContentControlAnalyzer**: Content controls
- **DirectFormatAnalyzer**: Direct formatting
- **InheritanceAnalyzer**: Style inheritance

### 3. Comparison Engine

```
Template Styles ──┐
                  ├── Style Comparator ──→ Mismatch Detection
Document Styles ──┘
                              ↓
                         Severity Assessment
                              ↓
                        Report Generation
```

#### Comparison Strategies
- **Strict Mode**: Exact property matching
- **Lenient Mode**: Semantic equivalence
- **Selective Comparison**: Configurable property filtering
- **Threshold-Based**: Acceptable variance ranges

## Data Architecture

### Entity Relationship Diagram

```
Template (1) ──────────── (*) TextStyle
    │                           │
    │                           │
    └── (*) VerificationResult  ├── (*) DirectFormatPattern
             │                  │
             │                  ├── (1) FormattingContext
             └── (*) Mismatch   │
                                └── (*) TabStop
```

### Entity Details

#### Template Entity
- **Primary Entity**: Represents document templates
- **Relationships**: One-to-many with TextStyle and VerificationResult
- **Business Rules**: Unique names, version control, audit trails
- **Storage**: File system path reference with hash verification

#### TextStyle Entity
- **Comprehensive Styling**: 100+ properties covering all Word formatting
- **Style Types**: Paragraph, Character, Table, List, Section
- **Normalization**: Consistent property representation
- **Signatures**: Hash-based fast comparison

#### VerificationResult Entity
- **Audit Trail**: Complete verification history
- **Relationships**: Many-to-one with Template, one-to-many with Mismatch
- **Status Tracking**: Pending, Completed, Failed states
- **Performance Metrics**: Processing time, mismatch counts

#### Mismatch Entity
- **Detailed Reporting**: Specific style differences
- **Context Information**: Location, sample text, structural role
- **Severity Classification**: Critical, High, Medium, Low
- **Actionable Insights**: Recommended fixes

### Database Design Principles

1. **Normalization**: 3NF compliance with controlled denormalization
2. **Indexing Strategy**: Query-optimized indexes on frequently accessed columns
3. **Audit Trails**: Created/Modified timestamps and user tracking
4. **Soft Deletes**: Status-based logical deletion
5. **Version Control**: Entity versioning for change tracking

### Index Strategy

```sql
-- Primary indexes for performance
CREATE INDEX idx_templates_status ON templates(status);
CREATE INDEX idx_templates_hash ON templates(file_hash);
CREATE INDEX idx_textstyles_template_signature ON textstyles(template_id, style_signature);
CREATE INDEX idx_textstyles_type ON textstyles(style_type);
CREATE INDEX idx_verificationresults_date ON verificationresults(verification_date);
CREATE INDEX idx_verificationresults_template ON verificationresults(template_id);
CREATE INDEX idx_mismatches_verification ON mismatches(verification_result_id);
CREATE INDEX idx_mismatches_severity ON mismatches(severity);

-- Composite indexes for complex queries
CREATE INDEX idx_templates_status_date ON templates(status, created_on);
CREATE INDEX idx_verification_template_date ON verificationresults(template_id, verification_date);
```

## Processing Pipeline

### 1. Template Creation Flow

```
1. File Upload & Validation
   ├── File format validation (.docx)
   ├── File size validation (< 100MB)
   ├── Virus scanning (future)
   └── Duplicate detection (hash)

2. Document Processing
   ├── OpenXML document opening
   ├── Schema validation
   ├── Corruption detection
   └── Security scanning

3. Style Extraction
   ├── Style definitions parsing
   ├── Effective style calculation
   ├── Inheritance resolution
   └── Direct formatting detection

4. Data Normalization
   ├── Property standardization
   ├── Unit conversions
   ├── Color normalization
   └── Signature generation

5. Database Storage
   ├── Transaction management
   ├── Entity relationships
   ├── Audit trail creation
   └── File system storage
```

### 2. Document Verification Flow

```
1. Request Processing
   ├── Document upload
   ├── Template resolution
   ├── Configuration parsing
   └── Validation

2. Style Extraction
   ├── Document parsing
   ├── Style extraction
   ├── Context analysis
   └── Normalization

3. Style Comparison
   ├── Template style loading
   ├── Matching algorithm
   ├── Difference detection
   └── Severity assessment

4. Report Generation
   ├── Mismatch categorization
   ├── Location identification
   ├── Sample text extraction
   └── Recommendation generation

5. Result Storage
   ├── Verification result creation
   ├── Mismatch detail storage
   ├── Performance metrics
   └── Response formatting
```

### 3. Style Extraction Algorithm

```python
# Pseudo-code for style extraction
def extract_styles(document):
    styles = []
    
    # Extract defined styles
    for style_def in document.styles:
        style = normalize_style(style_def)
        styles.append(style)
    
    # Extract effective styles
    for element in document.elements:
        effective_style = calculate_effective_style(element)
        if not matches_defined_style(effective_style, styles):
            direct_format = create_direct_format_pattern(effective_style)
            styles.append(direct_format)
    
    # Generate signatures
    for style in styles:
        style.signature = generate_signature(style)
    
    return styles

def calculate_effective_style(element):
    # Start with default style
    effective = get_default_style(element.type)
    
    # Apply style hierarchy
    if element.style_id:
        defined_style = get_defined_style(element.style_id)
        effective = merge_styles(effective, defined_style)
    
    # Apply direct formatting
    if element.direct_formatting:
        effective = merge_styles(effective, element.direct_formatting)
    
    return effective
```

## API Design

### RESTful Design Principles

1. **Resource-Based URLs**: Nouns representing entities
2. **HTTP Verbs**: Proper use of GET, POST, PUT, DELETE
3. **Status Codes**: Meaningful HTTP status codes
4. **Consistent Responses**: Uniform response structure
5. **Hypermedia**: HATEOAS-ready design (future enhancement)

### Endpoint Design Patterns

#### Collection Resources
```
GET    /api/templates              # List templates
POST   /api/templates              # Create template
```

#### Individual Resources
```
GET    /api/templates/{id}         # Get template
PUT    /api/templates/{id}         # Update template
DELETE /api/templates/{id}         # Delete template
```

#### Sub-Resources
```
GET    /api/templates/{id}/styles  # Get template styles
POST   /api/templates/{id}/process # Process template
GET    /api/templates/{id}/stats   # Get template stats
```

#### Action Resources
```
POST   /api/verification/verify    # Verify document
POST   /api/verification/compare   # Compare documents
GET    /api/verification/summary   # Get summary stats
```

### Request/Response Design

#### Consistent Response Wrapper
```json
{
  "data": {},
  "message": "Success",
  "timestamp": "2024-01-01T12:00:00Z",
  "requestId": "uuid"
}
```

#### Error Response Format
```json
{
  "error": {
    "code": "TEMPLATE_NOT_FOUND",
    "message": "Template with ID 123 not found",
    "details": {
      "templateId": 123,
      "requestPath": "/api/templates/123"
    }
  },
  "timestamp": "2024-01-01T12:00:00Z",
  "requestId": "uuid"
}
```

### Content Negotiation

- **Accept**: application/json (default), application/xml
- **Content-Type**: application/json, multipart/form-data
- **Compression**: gzip, deflate support
- **Versioning**: Header-based (X-API-Version)

## Security Architecture

### Authentication & Authorization

Current implementation is open access, but designed for easy integration:

```csharp
// Future authentication integration
builder.Services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options => {
        // JWT configuration
    });

builder.Services.AddAuthorization(options => {
    options.AddPolicy("TemplateManager", policy => 
        policy.RequireClaim("role", "template-manager"));
    options.AddPolicy("DocumentVerifier", policy => 
        policy.RequireClaim("role", "document-verifier"));
});
```

### Input Validation

Multi-layered validation approach:

1. **Model Validation**: Data annotations on DTOs
2. **Business Validation**: Service-level validation
3. **File Validation**: Format, size, content validation
4. **SQL Injection Prevention**: Parameterized queries through EF

### File Security

- **Upload Validation**: File type, size, content scanning
- **Secure Storage**: Isolated file system location
- **Access Control**: Authenticated access only (future)
- **Virus Scanning**: Integration ready for AV solutions

### Data Protection

- **Encryption at Rest**: Database and file system encryption
- **Encryption in Transit**: HTTPS/TLS for all communications
- **Sensitive Data**: No sensitive data in logs
- **GDPR Compliance**: Audit trails and data deletion capabilities

## Performance Considerations

### Database Optimization

#### Query Optimization
- **Eager Loading**: Explicit Include() statements
- **Projection**: Select only required columns
- **Pagination**: Limit result sets
- **Indexing**: Strategic index placement

#### Connection Management
- **Connection Pooling**: Configured pool sizes
- **Connection Lifetime**: Appropriate timeout settings
- **Bulk Operations**: Batch processing for large datasets

### Memory Management

#### Document Processing
- **Streaming**: Large document processing
- **Disposal**: Proper resource cleanup
- **Caching**: Template style caching
- **Memory Limits**: Configurable memory boundaries

#### Garbage Collection
- **Large Object Heap**: Minimal LOH allocations
- **Generation Awareness**: Short-lived object patterns
- **Memory Pressure**: Monitoring and alerting

### I/O Optimization

#### File System
- **Async Operations**: Non-blocking file operations
- **Buffering**: Appropriate buffer sizes
- **SSD Optimization**: Sequential access patterns
- **Cleanup**: Temporary file management

#### Network I/O
- **Response Compression**: gzip compression
- **Chunked Transfer**: Large response handling
- **Keep-Alive**: Connection reuse
- **Timeout Management**: Appropriate timeout settings

## Scalability Design

### Horizontal Scaling

#### Stateless Design
- **No Session State**: Stateless API design
- **Load Balancer Ready**: Health check endpoints
- **Shared Storage**: Database and file system sharing
- **Configuration**: Environment-based configuration

#### Database Scaling
- **Read Replicas**: Read-only database instances
- **Connection Pooling**: Per-instance pool management
- **Partitioning**: Table partitioning strategies
- **Caching**: Redis integration ready

### Vertical Scaling

#### Resource Utilization
- **CPU Intensive**: Document processing optimization
- **Memory Management**: Efficient memory usage
- **I/O Optimization**: Disk and network optimization
- **Threading**: Async/await patterns

#### Performance Monitoring
- **Metrics Collection**: Application metrics
- **Health Checks**: System health endpoints
- **Alerting**: Performance threshold alerts
- **Profiling**: Performance profiling capabilities

### Microservices Readiness

Current monolithic design can be decomposed into:

1. **Template Service**: Template management
2. **Document Service**: Document processing
3. **Verification Service**: Style comparison
4. **Reporting Service**: Report generation
5. **File Service**: File storage and management

## Integration Points

### External Systems

#### Future Integrations
- **Active Directory**: User authentication
- **SharePoint**: Document repository integration
- **Office 365**: Online document processing
- **Workflow Systems**: Process automation
- **Notification Systems**: Email/SMS notifications

#### API Gateway Integration
- **Rate Limiting**: Request throttling
- **Authentication**: Centralized auth
- **Logging**: Centralized logging
- **Monitoring**: Request monitoring

### Webhook Support

Future webhook implementation:

```csharp
public class WebhookService
{
    public async Task NotifyVerificationComplete(VerificationResult result)
    {
        var payload = new {
            EventType = "verification.completed",
            Data = result,
            Timestamp = DateTime.UtcNow
        };
        
        await httpClient.PostAsJsonAsync(webhookUrl, payload);
    }
}
```

### Message Queue Integration

Ready for async processing:

```csharp
// Future message queue integration
public interface IMessagePublisher
{
    Task PublishAsync<T>(T message, string queue);
}

public class DocumentProcessingService
{
    public async Task ProcessTemplateAsync(int templateId)
    {
        // Process template
        await _messagePublisher.PublishAsync(
            new TemplateProcessedEvent { TemplateId = templateId },
            "template-processed"
        );
    }
}
```

This technical architecture provides a solid foundation for the DocStyleVerify.API system, with clear separation of concerns, scalability considerations, and extensibility for future enhancements. 