# API Reference

Complete reference documentation for DocStyleVerify.API endpoints.

## Base URL
```
https://localhost:5001/api
```

## Response Format
All API responses follow a consistent format:

### Success Response
```json
{
  "data": [response_data],
  "message": "Success message",
  "timestamp": "2024-01-01T12:00:00Z"
}
```

### Error Response
```json
{
  "error": {
    "code": "ERROR_CODE",
    "message": "Error description",
    "details": "Additional error details"
  },
  "timestamp": "2024-01-01T12:00:00Z"
}
```

## HTTP Status Codes
- `200 OK` - Request successful
- `201 Created` - Resource created successfully
- `204 No Content` - Request successful with no content
- `400 Bad Request` - Invalid request parameters
- `404 Not Found` - Resource not found
- `409 Conflict` - Resource conflict
- `500 Internal Server Error` - Server error

---

## Templates API

### Get All Templates
**GET** `/templates`

Retrieves a paginated list of templates with optional filtering.

#### Query Parameters
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `page` | integer | 1 | Page number (minimum: 1) |
| `pageSize` | integer | 20 | Items per page (1-100) |
| `status` | string | - | Filter by status (Active, Inactive, Archived) |

#### Response
```json
{
  "data": [
    {
      "id": 1,
      "name": "Corporate Template",
      "description": "Standard corporate document template",
      "fileName": "template.docx",
      "filePath": "Templates/corporate.docx",
      "fileHash": "abc123def456",
      "fileSize": 1024000,
      "status": "Active",
      "createdBy": "admin@company.com",
      "createdOn": "2024-01-01T10:00:00Z",
      "modifiedBy": "",
      "modifiedOn": null,
      "version": 1,
      "textStylesCount": 25
    }
  ]
}
```

### Get Template by ID
**GET** `/templates/{id}`

Retrieves a specific template by its ID.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Template ID |

#### Response
```json
{
  "data": {
    "id": 1,
    "name": "Corporate Template",
    "description": "Standard corporate document template",
    "fileName": "template.docx",
    "filePath": "Templates/corporate.docx",
    "fileHash": "abc123def456",
    "fileSize": 1024000,
    "status": "Active",
    "createdBy": "admin@company.com",
    "createdOn": "2024-01-01T10:00:00Z",
    "modifiedBy": "",
    "modifiedOn": null,
    "version": 1,
    "textStylesCount": 25
  }
}
```

### Create Template
**POST** `/templates`

Creates a new template from an uploaded Word document.

#### Request Body (multipart/form-data)
| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `name` | string | Yes | Template name (max 200 chars) |
| `description` | string | No | Template description (max 500 chars) |
| `file` | file | Yes | Word document file (.docx) |
| `createdBy` | string | Yes | User email (max 100 chars) |

#### Response (201 Created)
```json
{
  "data": {
    "id": 2,
    "name": "New Template",
    "description": "Description",
    "fileName": "new-template.docx",
    "filePath": "Templates/new-template.docx",
    "fileHash": "def456ghi789",
    "fileSize": 2048000,
    "status": "Active",
    "createdBy": "user@company.com",
    "createdOn": "2024-01-01T11:00:00Z",
    "modifiedBy": "",
    "modifiedOn": null,
    "version": 1,
    "textStylesCount": 0
  }
}
```

### Update Template
**PUT** `/templates/{id}`

Updates an existing template's metadata.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Template ID |

#### Request Body (application/json)
```json
{
  "name": "Updated Template Name",
  "description": "Updated description",
  "status": "Active",
  "modifiedBy": "user@company.com"
}
```

#### Response
```json
{
  "data": {
    "id": 1,
    "name": "Updated Template Name",
    "description": "Updated description",
    "fileName": "template.docx",
    "filePath": "Templates/corporate.docx",
    "fileHash": "abc123def456",
    "fileSize": 1024000,
    "status": "Active",
    "createdBy": "admin@company.com",
    "createdOn": "2024-01-01T10:00:00Z",
    "modifiedBy": "user@company.com",
    "modifiedOn": "2024-01-01T12:00:00Z",
    "version": 2,
    "textStylesCount": 25
  }
}
```

### Delete Template
**DELETE** `/templates/{id}`

Deletes a template and all associated data.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Template ID |

#### Response (204 No Content)
No response body.

### Process Template
**POST** `/templates/{id}/process`

Processes a template to extract all style information.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Template ID |

#### Response
```json
{
  "data": {
    "templateId": 1,
    "templateName": "Corporate Template",
    "totalTextStyles": 25,
    "totalDirectFormatPatterns": 5,
    "extractedStyleTypes": ["Paragraph", "Character", "Table", "List"],
    "processingMessages": [
      "Successfully extracted 25 text styles",
      "Found 5 direct format patterns",
      "Processing completed in 2.5 seconds"
    ],
    "success": true,
    "errorMessage": "",
    "processedOn": "2024-01-01T12:00:00Z"
  }
}
```

### Get Template Styles
**GET** `/templates/{id}/styles`

Retrieves all styles for a specific template.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Template ID |

#### Query Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `styleType` | string | Filter by style type (Paragraph, Character, Table, List, Section) |

#### Response
```json
{
  "data": [
    {
      "id": 1,
      "name": "Normal",
      "fontFamily": "Calibri",
      "fontSize": 11.0,
      "isBold": false,
      "isItalic": false,
      "isUnderline": false,
      "color": "#000000",
      "alignment": "Left",
      "spacingBefore": 0.0,
      "spacingAfter": 0.0,
      "indentationLeft": 0.0,
      "indentationRight": 0.0,
      "lineSpacing": 1.0,
      "styleType": "Paragraph",
      "styleSignature": "Normal|Calibri|11|False|False|#000000|Left"
    }
  ]
}
```

### Get Template Statistics
**GET** `/templates/{id}/stats`

Retrieves statistics and metadata for a template.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Template ID |

#### Response
```json
{
  "data": {
    "templateId": 1,
    "templateName": "Corporate Template",
    "totalStyles": 25,
    "styleTypeBreakdown": {
      "Paragraph": 15,
      "Character": 5,
      "Table": 3,
      "List": 2
    },
    "directFormatPatternsCount": 5,
    "lastProcessed": "2024-01-01T12:00:00Z",
    "fileSize": 1024000,
    "status": "Active"
  }
}
```

### Validate Template File
**POST** `/templates/validate`

Validates a Word document file before creating a template.

#### Request Body (multipart/form-data)
| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `file` | file | Yes | Word document file (.docx) |

#### Response
```json
{
  "data": {
    "fileName": "template.docx",
    "fileSize": 1024000,
    "isValid": true,
    "validationMessages": [
      "File format is valid",
      "File size is within limits",
      "Document structure is valid"
    ]
  }
}
```

---

## Verification API

### Verify Document
**POST** `/verification/verify`

Verifies a document against a template and returns mismatches.

#### Request Body (multipart/form-data)
| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `document` | file | Yes | Word document file (.docx) |
| `templateId` | integer | Yes | Template ID to verify against |
| `createdBy` | string | Yes | User email |
| `strictMode` | boolean | No | Enable strict comparison (default: true) |
| `ignoreStyleTypes` | array | No | Style types to ignore during verification |

#### Response
```json
{
  "data": {
    "id": 1,
    "templateId": 1,
    "templateName": "Corporate Template",
    "documentName": "document.docx",
    "verificationDate": "2024-01-01T13:00:00Z",
    "totalMismatches": 3,
    "status": "Completed",
    "createdBy": "user@company.com",
    "mismatches": [
      {
        "id": 1,
        "contextKey": "paragraph_1",
        "location": "Page 1, Paragraph 1",
        "structuralRole": "Heading",
        "expectedStyle": {
          "fontFamily": "Arial",
          "fontSize": 14.0,
          "isBold": true
        },
        "actualStyle": {
          "fontFamily": "Calibri",
          "fontSize": 12.0,
          "isBold": false
        },
        "mismatchFields": ["FontFamily", "FontSize", "IsBold"],
        "sampleText": "Chapter 1: Introduction",
        "severity": "High"
      }
    ]
  }
}
```

### Get Verification Results
**GET** `/verification/results`

Retrieves verification results with optional filtering.

#### Query Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `templateId` | integer | Filter by template ID |
| `page` | integer | Page number (default: 1) |
| `pageSize` | integer | Items per page (1-100, default: 20) |
| `status` | string | Filter by status (Pending, Completed, Failed) |

#### Response
```json
{
  "data": [
    {
      "id": 1,
      "templateId": 1,
      "templateName": "Corporate Template",
      "documentName": "document.docx",
      "verificationDate": "2024-01-01T13:00:00Z",
      "totalMismatches": 3,
      "status": "Completed",
      "createdBy": "user@company.com"
    }
  ]
}
```

### Get Verification Result by ID
**GET** `/verification/results/{id}`

Retrieves a specific verification result with detailed mismatches.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Verification result ID |

#### Response
```json
{
  "data": {
    "id": 1,
    "templateId": 1,
    "templateName": "Corporate Template",
    "documentName": "document.docx",
    "verificationDate": "2024-01-01T13:00:00Z",
    "totalMismatches": 3,
    "status": "Completed",
    "createdBy": "user@company.com",
    "mismatches": [
      {
        "id": 1,
        "contextKey": "paragraph_1",
        "location": "Page 1, Paragraph 1",
        "structuralRole": "Heading",
        "expectedStyle": {
          "fontFamily": "Arial",
          "fontSize": 14.0,
          "isBold": true
        },
        "actualStyle": {
          "fontFamily": "Calibri",
          "fontSize": 12.0,
          "isBold": false
        },
        "mismatchFields": ["FontFamily", "FontSize", "IsBold"],
        "sampleText": "Chapter 1: Introduction",
        "severity": "High"
      }
    ]
  }
}
```

### Get Mismatch Report
**GET** `/verification/results/{id}/report`

Generates a detailed mismatch report for a verification result.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Verification result ID |

#### Response
```json
{
  "data": [
    {
      "id": 1,
      "contextKey": "paragraph_1",
      "location": "Page 1, Paragraph 1",
      "structuralRole": "Heading",
      "mismatchType": "Style Mismatch",
      "description": "Font family, size, and bold formatting do not match template",
      "expectedValue": "Arial, 14pt, Bold",
      "actualValue": "Calibri, 12pt, Regular",
      "severity": "High",
      "sampleText": "Chapter 1: Introduction",
      "recommendedAction": "Change font to Arial, increase size to 14pt, and apply bold formatting",
      "affectedElements": 1
    }
  ]
}
```

### Get Verification Summary
**GET** `/verification/summary`

Retrieves verification statistics and summary information.

#### Query Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `templateId` | integer | Filter by template ID |
| `fromDate` | date | Start date filter (YYYY-MM-DD) |
| `toDate` | date | End date filter (YYYY-MM-DD) |

#### Response
```json
{
  "data": {
    "totalVerifications": 100,
    "successfulVerifications": 85,
    "failedVerifications": 15,
    "averageMismatchesPerDocument": 2.5,
    "mostCommonMismatches": [
      {
        "mismatchType": "FontFamily",
        "count": 25,
        "percentage": 35.7
      },
      {
        "mismatchType": "FontSize",
        "count": 18,
        "percentage": 25.7
      }
    ],
    "severityBreakdown": {
      "Critical": 5,
      "High": 15,
      "Medium": 35,
      "Low": 20
    },
    "templateUsage": [
      {
        "templateId": 1,
        "templateName": "Corporate Template",
        "verificationCount": 60,
        "averageMismatches": 2.1
      }
    ]
  }
}
```

### Compare Documents
**POST** `/verification/compare`

Compares two documents and identifies style differences.

#### Request Body (multipart/form-data)
| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `document1` | file | Yes | First Word document file (.docx) |
| `document2` | file | Yes | Second Word document file (.docx) |
| `createdBy` | string | Yes | User email |

#### Response
```json
{
  "data": [
    {
      "contextKey": "paragraph_1",
      "location": "Page 1, Paragraph 1",
      "document1Style": {
        "fontFamily": "Arial",
        "fontSize": 12.0,
        "isBold": true
      },
      "document2Style": {
        "fontFamily": "Calibri",
        "fontSize": 11.0,
        "isBold": false
      },
      "differences": ["FontFamily", "FontSize", "IsBold"],
      "similarity": 0.65
    }
  ]
}
```

### Delete Verification Result
**DELETE** `/verification/results/{id}`

Deletes a verification result and all associated mismatches.

#### Path Parameters
| Parameter | Type | Description |
|-----------|------|-------------|
| `id` | integer | Verification result ID |

#### Response (204 No Content)
No response body.

---

## Error Codes

### Template Errors
| Code | Message | Description |
|------|---------|-------------|
| `TEMPLATE_NOT_FOUND` | Template not found | Template with specified ID does not exist |
| `TEMPLATE_NAME_EXISTS` | Template name already exists | A template with the same name already exists |
| `INVALID_TEMPLATE_FILE` | Invalid template file | Uploaded file is not a valid .docx document |
| `TEMPLATE_PROCESSING_FAILED` | Template processing failed | Error occurred while processing template |

### Verification Errors
| Code | Message | Description |
|------|---------|-------------|
| `VERIFICATION_NOT_FOUND` | Verification result not found | Verification result with specified ID does not exist |
| `INVALID_DOCUMENT_FILE` | Invalid document file | Uploaded file is not a valid .docx document |
| `VERIFICATION_FAILED` | Document verification failed | Error occurred during document verification |
| `TEMPLATE_NOT_PROCESSED` | Template not processed | Template must be processed before verification |

### General Errors
| Code | Message | Description |
|------|---------|-------------|
| `INVALID_REQUEST` | Invalid request | Request parameters are invalid |
| `FILE_TOO_LARGE` | File too large | Uploaded file exceeds size limit |
| `UNSUPPORTED_FORMAT` | Unsupported file format | File format is not supported |
| `PROCESSING_TIMEOUT` | Processing timeout | Operation timed out |
| `INTERNAL_ERROR` | Internal server error | Unexpected server error occurred |

---

## Rate Limiting

The API implements rate limiting to prevent abuse:
- **Standard endpoints**: 100 requests per minute per IP
- **File upload endpoints**: 10 requests per minute per IP
- **Processing endpoints**: 5 requests per minute per IP

Rate limit headers are included in all responses:
```
X-RateLimit-Limit: 100
X-RateLimit-Remaining: 95
X-RateLimit-Reset: 1640995200
```

## Authentication

Currently, the API does not require authentication, but it's designed to easily integrate with various authentication providers. All endpoints accept an optional `Authorization` header for future compatibility.

## File Upload Limits

- **Maximum file size**: 100 MB
- **Supported formats**: .docx only
- **Content-Type**: multipart/form-data
- **Timeout**: 10 minutes for processing

## Pagination

List endpoints support pagination with the following parameters:
- `page`: Page number (starting from 1)
- `pageSize`: Number of items per page (1-100)

Pagination information is included in response headers:
```
X-Pagination-Page: 1
X-Pagination-PageSize: 20
X-Pagination-TotalPages: 5
X-Pagination-TotalItems: 100
``` 