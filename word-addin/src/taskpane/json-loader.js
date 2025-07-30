/**
 * JSON Loader Utility
 * Handles loading suggestions from various sources
 */

/**
 * Load suggestions from the actual JSON file in the parent project
 */
export async function loadActualSuggestions() {
    try {
        // Try to load from the parent project's Extracted_JSON folder
        const response = await fetch('../../Extracted_JSON/updated_response_json.json');
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const jsonData = await response.json();
        return jsonData;
    } catch (error) {
        console.warn('Could not load actual JSON file:', error);
        
        // Fallback to demo data
        return getDemoSuggestions();
    }
}

/**
 * Load suggestions from a file input
 */
export function loadFromFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const jsonData = JSON.parse(e.target.result);
                resolve(jsonData);
            } catch (error) {
                reject(new Error('Invalid JSON file: ' + error.message));
            }
        };
        
        reader.onerror = function() {
            reject(new Error('Error reading file'));
        };
        
        reader.readAsText(file);
    });
}

/**
 * Load suggestions from API endpoint
 */
export async function loadFromAPI(endpoint) {
    try {
        const response = await fetch(endpoint, {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
            }
        });
        
        if (!response.ok) {
            throw new Error(`API error! status: ${response.status}`);
        }
        
        return await response.json();
    } catch (error) {
        throw new Error('Failed to load from API: ' + error.message);
    }
}

/**
 * Get demo suggestions for testing
 */
export function getDemoSuggestions() {
    return {
        "suggestions": {
            "document": [
                {
                    "json_object": [
                        {
                            "id": 277,
                            "templateId": 2,
                            "name": "Heading 4_36",
                            "fontFamily": "",
                            "fontSize": 0,
                            "isBold": true,
                            "isItalic": true,
                            "isUnderline": true,
                            "isStrikethrough": false,
                            "isAllCaps": false,
                            "isSmallCaps": false,
                            "hasShadow": false,
                            "hasOutline": false,
                            "highlighting": "",
                            "characterSpacing": 0,
                            "language": "",
                            "color": "4F81BD",
                            "alignment": "",
                            "spacingBefore": 10,
                            "spacingAfter": 0,
                            "indentationLeft": 0,
                            "indentationRight": 0,
                            "firstLineIndent": 0,
                            "widowOrphanControl": false,
                            "keepWithNext": false,
                            "keepTogether": false,
                            "tabStops": [],
                            "borderStyle": "",
                            "borderColor": "",
                            "borderWidth": 0,
                            "borderDirections": "",
                            "listNumberStyle": "",
                            "listLevel": 0,
                            "lineSpacing": 1,
                            "styleType": "Heading",
                            "basedOnStyle": "Heading4",
                            "priority": 0,
                            "styleSignature": "GQKPohtArgSFPsqmTIohyiSeIJG/s5HR",
                            "formattingContext": {
                                "id": 277,
                                "styleName": "",
                                "elementType": "Paragraph",
                                "parentContext": "",
                                "structuralRole": "Heading",
                                "sampleText": "14. Table of Contents (Field)",
                                "contentControlTag": "",
                                "contentControlType": "",
                                "sectionIndex": 0,
                                "tableIndex": -1,
                                "rowIndex": -1,
                                "cellIndex": -1,
                                "paragraphIndex": 36,
                                "runIndex": -1,
                                "contextKey": "Section:0:Paragraph:36",
                                "contentControlProperties": {}
                            },
                            "directFormatPatterns": [],
                            "createdBy": "System",
                            "createdOn": "2025-07-18T12:16:23.840968Z",
                            "modifiedBy": "",
                            "modifiedOn": null,
                            "version": 1
                        }
                    ],
                    "ops": [
                        {
                            "prop": "paragraph.style",
                            "to": "Heading2"
                        }
                    ],
                    "message": "In the target document, Heading 4 is used where Heading 2 should be"
                },
                {
                    "json_object": [
                        {
                            "id": 266,
                            "templateId": 2,
                            "name": "ListBullet_16",
                            "styleType": "List",
                            "basedOnStyle": "ListBullet",
                            "formattingContext": {
                                "id": 266,
                                "sampleText": "First step",
                                "paragraphIndex": 16,
                                "contextKey": "Section:0:Paragraph:16"
                            }
                        },
                        {
                            "id": 267,
                            "templateId": 2,
                            "name": "ListBullet_17",
                            "styleType": "List",
                            "basedOnStyle": "ListBullet",
                            "formattingContext": {
                                "id": 267,
                                "sampleText": "Second step",
                                "paragraphIndex": 17,
                                "contextKey": "Section:0:Paragraph:17"
                            }
                        },
                        {
                            "id": 268,
                            "templateId": 2,
                            "name": "ListBullet_18",
                            "styleType": "List",
                            "basedOnStyle": "ListBullet",
                            "formattingContext": {
                                "id": 268,
                                "sampleText": "Third step",
                                "paragraphIndex": 18,
                                "contextKey": "Section:0:Paragraph:18"
                            }
                        }
                    ],
                    "ops": [
                        {
                            "prop": "paragraph.style",
                            "to": "ListNumber"
                        }
                    ],
                    "message": "template has a Bullet list and Numbered list under section 3, check that list logic did not break"
                },
                {
                    "json_object": [
                        {
                            "id": 275,
                            "templateId": 2,
                            "name": "Paragraph_32_7840",
                            "fontFamily": "Comic Sans MS",
                            "styleType": "Paragraph",
                            "formattingContext": {
                                "id": 275,
                                "sampleText": "Content before page break.",
                                "paragraphIndex": 32,
                                "contextKey": "Section:0:Paragraph:32"
                            },
                            "directFormatPatterns": [
                                {
                                    "id": 30,
                                    "textStyleId": 275,
                                    "patternName": "Run_32_0",
                                    "context": "Paragraph:32,Run:0",
                                    "fontFamily": "Comic Sans MS",
                                    "sampleText": "Content before page break.",
                                    "occurrenceCount": 1
                                }
                            ]
                        }
                    ],
                    "ops": [
                        {
                            "prop": "font.name",
                            "to": "Cambria"
                        }
                    ],
                    "message": "template has the direct format font Cambria, which is changed to Comic Sans MS in the current document"
                }
            ],
            "styles": [
                {
                    "styleId": "Heading2",
                    "ops": [
                        {
                            "prop": "style.basedOn",
                            "to": "Normal"
                        },
                        {
                            "prop": "style.font.color",
                            "to": "4F81BD"
                        }
                    ],
                    "message": "In the template, Heading 2 is based on Normal, in the current document that was changed to Body Text. The font colour for Heading 2 was changed from 4F81BD to F79646"
                },
                {
                    "styleId": "illegalstyle",
                    "ops": [
                        {
                            "prop": "style.isHidden",
                            "to": true
                        }
                    ],
                    "message": "illegal_style does not appear in template, consider deleting it"
                }
            ]
        }
    };
}

/**
 * Validate suggestions JSON structure
 */
export function validateSuggestions(jsonData) {
    if (!jsonData) {
        throw new Error('No data provided');
    }
    
    if (!jsonData.suggestions) {
        throw new Error('Missing "suggestions" property');
    }
    
    if (!jsonData.suggestions.document || !Array.isArray(jsonData.suggestions.document)) {
        throw new Error('Missing or invalid "suggestions.document" array');
    }
    
    // Validate each suggestion
    jsonData.suggestions.document.forEach((suggestion, index) => {
        if (!suggestion.json_object || !Array.isArray(suggestion.json_object)) {
            throw new Error(`Suggestion ${index}: Missing or invalid "json_object" array`);
        }
        
        if (!suggestion.ops || !Array.isArray(suggestion.ops)) {
            throw new Error(`Suggestion ${index}: Missing or invalid "ops" array`);
        }
        
        if (!suggestion.message) {
            throw new Error(`Suggestion ${index}: Missing "message" property`);
        }
        
        // Validate each object in json_object
        suggestion.json_object.forEach((obj, objIndex) => {
            if (!obj.formattingContext) {
                throw new Error(`Suggestion ${index}, Object ${objIndex}: Missing "formattingContext"`);
            }
            
            if (typeof obj.formattingContext.paragraphIndex !== 'number') {
                throw new Error(`Suggestion ${index}, Object ${objIndex}: Invalid "paragraphIndex"`);
            }
        });
        
        // Validate operations
        suggestion.ops.forEach((op, opIndex) => {
            if (!op.prop || !op.to) {
                throw new Error(`Suggestion ${index}, Operation ${opIndex}: Missing "prop" or "to" property`);
            }
        });
    });
    
    return true;
} 