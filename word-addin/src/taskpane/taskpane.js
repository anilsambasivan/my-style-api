/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("loadSuggestionsBtn").onclick = loadSuggestions;
        document.getElementById("applyButton").onclick = applyCurrentSuggestion;
        document.getElementById("skipButton").onclick = skipCurrentSuggestion;
        
        // Initialize the add-in
        initializeAddin();
    }
});

// Global state
let suggestions = null;
let currentSuggestionIndex = 0;
let processedSuggestions = [];

/**
 * Initialize the add-in
 */
function initializeAddin() {
    showStatus("Welcome! Load your style suggestions JSON file to get started.", "info");
}

/**
 * Load suggestions from JSON file
 */
function loadSuggestions() {
    const fileInput = document.getElementById("fileInput");
    
    // For demo purposes, we'll use the provided JSON structure
    // In production, this would load from a file or API
    loadDemoSuggestions();
    
    // Uncomment this for file loading functionality:
    // fileInput.click();
    // fileInput.onchange = handleFileLoad;
}

/**
 * Load demo suggestions (using the actual JSON file)
 */
async function loadDemoSuggestions() {
    try {
        // Try to load from the actual JSON file
        const response = await fetch('../updated_response_json.json');
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const jsonData = await response.json();
        processSuggestions(jsonData);
        
    } catch (error) {
        console.warn('Could not load actual JSON file:', error);
        showStatus("Could not load JSON file. Please check if updated_response_json.json is accessible.", "error");
        
        // Fallback to the demo data structure based on your JSON
        const fallbackSuggestions = {
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
                                }
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
                    }
                ]
            }
        };
        
        processSuggestions(fallbackSuggestions);
    }
}

/**
 * Handle file load from input
 */
function handleFileLoad(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const jsonData = JSON.parse(e.target.result);
            processSuggestions(jsonData);
        } catch (error) {
            showStatus("Error reading JSON file: " + error.message, "error");
        }
    };
    reader.readAsText(file);
}

/**
 * Process the loaded suggestions
 */
function processSuggestions(jsonData) {
    if (!jsonData.suggestions || !jsonData.suggestions.document) {
        showStatus("Invalid JSON format. Please ensure the file contains 'suggestions.document' array.", "error");
        return;
    }

    suggestions = jsonData.suggestions.document;
    currentSuggestionIndex = 0;
    processedSuggestions = [];

    if (suggestions.length === 0) {
        showStatus("No suggestions found in the file.", "info");
        return;
    }

    showStatus(`Loaded ${suggestions.length} suggestions successfully!`, "success");
    
    // Hide load section and show main content
    document.getElementById("loadSection").classList.add("hidden");
    document.getElementById("mainContent").classList.remove("hidden");
    document.getElementById("progressContainer").classList.remove("hidden");

    // Display first suggestion
    displayCurrentSuggestion();
}

/**
 * Display the current suggestion
 */
function displayCurrentSuggestion() {
    if (currentSuggestionIndex >= suggestions.length) {
        showCompletionState();
        return;
    }

    const suggestion = suggestions[currentSuggestionIndex];
    const mainObject = suggestion.json_object[0];
    const context = mainObject.formattingContext;

    // Update progress
    updateProgress();

    // Update suggestion card
    document.getElementById("suggestionTitle").textContent = 
        `Suggestion ${currentSuggestionIndex + 1} of ${suggestions.length}`;
    
    document.getElementById("suggestionLocation").textContent = 
        `Paragraph ${context.paragraphIndex + 1}${context.sectionIndex > 0 ? `, Section ${context.sectionIndex + 1}` : ''}`;
    
    document.getElementById("suggestionMessage").textContent = suggestion.message;

    // Update previews
    updatePreviews(suggestion, mainObject);

    // Navigate to the paragraph in Word
    navigateToTarget(context.paragraphIndex);
}

/**
 * Update the progress bar and text
 */
function updateProgress() {
    const progress = ((currentSuggestionIndex) / suggestions.length) * 100;
    document.getElementById("progressFill").style.width = `${progress}%`;
    document.getElementById("progressText").textContent = 
        `Processing suggestion ${currentSuggestionIndex + 1} of ${suggestions.length}`;
}

/**
 * Update the preview boxes
 */
function updatePreviews(suggestion, mainObject) {
    const context = mainObject.formattingContext;
    
    // Current style preview
    const currentPreview = document.getElementById("currentPreview");
    let currentStyleText = `"${context.sampleText}"`;
    
    // Add current formatting info
    if (mainObject.fontFamily) {
        currentStyleText += `\n\nFont: ${mainObject.fontFamily}`;
    }
    if (mainObject.basedOnStyle) {
        currentStyleText += `\nStyle: ${mainObject.basedOnStyle}`;
    }
    if (mainObject.color) {
        currentStyleText += `\nColor: #${mainObject.color}`;
    }
    
    currentPreview.textContent = currentStyleText;

    // Suggested style preview
    const suggestedPreview = document.getElementById("suggestedPreview");
    let suggestedStyleText = `"${context.sampleText}"`;
    
    // Add suggested changes
    suggestion.ops.forEach(op => {
        if (op.prop === "paragraph.style") {
            suggestedStyleText += `\n\nNew Style: ${op.to}`;
        } else if (op.prop === "font.name") {
            suggestedStyleText += `\n\nNew Font: ${op.to}`;
        } else if (op.prop === "style.font.color") {
            suggestedStyleText += `\n\nNew Color: ${op.to}`;
        }
    });
    
    suggestedPreview.textContent = suggestedStyleText;
}

/**
 * Navigate to the target paragraph in Word
 */
async function navigateToTarget(paragraphIndex) {
    try {
        await Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("length");
            await context.sync();
            
            if (paragraphIndex < paragraphs.items.length) {
                const targetParagraph = paragraphs.items[paragraphIndex];
                targetParagraph.select();
                await context.sync();
                
                showStatus("Navigated to target paragraph", "info");
            } else {
                showStatus(`Warning: Paragraph ${paragraphIndex + 1} not found in document`, "error");
            }
        });
    } catch (error) {
        console.error("Error navigating to paragraph:", error);
        showStatus("Error navigating to paragraph: " + error.message, "error");
    }
}

/**
 * Apply the current suggestion
 */
async function applyCurrentSuggestion() {
    if (currentSuggestionIndex >= suggestions.length) return;

    const suggestion = suggestions[currentSuggestionIndex];
    const mainObject = suggestion.json_object[0];
    const context = mainObject.formattingContext;

    // Disable buttons during processing
    setButtonsEnabled(false);
    showStatus("Applying suggestion...", "info");

    try {
        await Word.run(async (wordContext) => {
            const paragraphs = wordContext.document.body.paragraphs;
            paragraphs.load("items");
            await wordContext.sync();
            
            console.log(`Document has ${paragraphs.items.length} paragraphs`);
            
            // Apply changes to all affected paragraphs
            for (const obj of suggestion.json_object) {
                const paragraphIndex = obj.formattingContext.paragraphIndex;
                
                if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
                    console.warn(`Paragraph index ${paragraphIndex} is out of range (document has ${paragraphs.items.length} paragraphs)`);
                    continue;
                }
                
                const targetParagraph = paragraphs.items[paragraphIndex];
                console.log(`Applying changes to paragraph ${paragraphIndex}: "${obj.formattingContext.sampleText}"`);
                
                // Apply each operation individually with error handling
                for (const op of suggestion.ops) {
                    try {
                        await applyOperation(targetParagraph, op);
                        await wordContext.sync(); // Sync after each operation
                        console.log(`Successfully applied: ${op.prop} = ${op.to}`);
                    } catch (opError) {
                        console.error(`Failed to apply operation ${op.prop} = ${op.to}:`, opError);
                        showStatus(`Warning: Could not apply ${op.prop} change. Continuing with next operation.`, "error");
                        // Continue with other operations instead of failing completely
                    }
                }
            }
            
            await wordContext.sync();
            
            // Mark as processed
            processedSuggestions.push({
                index: currentSuggestionIndex,
                suggestion: suggestion,
                applied: true,
                timestamp: new Date()
            });
            
            showStatus("Suggestion applied successfully!", "success");
            
            // Move to next suggestion
            setTimeout(() => {
                moveToNextSuggestion();
            }, 1000);
        });
    } catch (error) {
        console.error("Error applying suggestion:", error);
        showStatus("Error applying suggestion: " + error.message, "error");
        setButtonsEnabled(true);
    }
}



/**
 * Apply a specific operation to a paragraph
 */
async function applyOperation(paragraph, operation) {
    try {
        console.log(`Applying operation: ${operation.prop} = ${operation.to}`);
        
        switch (operation.prop) {
            case "paragraph.style":
                if (operation.to && operation.to.trim() !== "") {
                    // Map common style names to ensure compatibility
                    const styleMap = {
                        "Heading2": "Heading 2",
                        "Heading1": "Heading 1", 
                        "Heading3": "Heading 3",
                        "Heading4": "Heading 4",
                        "ListNumber": "List Number",
                        "ListBullet": "List Bullet",
                        "ListParagraph": "List Paragraph",
                        "Normal": "Normal",
                        "BodyText": "Body Text"
                    };
                    
                    const styleName = styleMap[operation.to] || operation.to;
                    console.log(`Setting paragraph style to: "${styleName}"`);
                    paragraph.style = styleName;
                } else {
                    console.warn("Invalid style name:", operation.to);
                }
                break;
                
            case "font.name":
                if (operation.to && operation.to.trim() !== "") {
                    paragraph.font.name = operation.to;
                } else {
                    console.warn("Invalid font name:", operation.to);
                }
                break;
                
            case "style.font.color":
                if (operation.to && operation.to.trim() !== "") {
                    // Ensure color format is correct (add # if missing)
                    let color = operation.to;
                    if (!color.startsWith("#") && color.length === 6) {
                        color = "#" + color;
                    }
                    paragraph.font.color = color;
                } else {
                    console.warn("Invalid color value:", operation.to);
                }
                break;
                
            case "paragraph.alignment":
                if (operation.to && operation.to.trim() !== "") {
                    // Map alignment values to Word constants
                    const alignmentMap = {
                        "Left": Word.Alignment.left,
                        "Center": Word.Alignment.centered, 
                        "Right": Word.Alignment.right,
                        "Justify": Word.Alignment.justified
                    };
                    paragraph.alignment = alignmentMap[operation.to] || operation.to;
                } else {
                    console.warn("Invalid alignment value:", operation.to);
                }
                break;
                
            default:
                console.warn("Unknown operation:", operation.prop);
        }
    } catch (error) {
        console.error(`Error applying operation ${operation.prop}:`, error);
        throw error;
    }
}

/**
 * Skip the current suggestion
 */
function skipCurrentSuggestion() {
    const suggestion = suggestions[currentSuggestionIndex];
    
    // Mark as processed but not applied
    processedSuggestions.push({
        index: currentSuggestionIndex,
        suggestion: suggestion,
        applied: false,
        skipped: true,
        timestamp: new Date()
    });
    
    showStatus("Suggestion skipped", "info");
    moveToNextSuggestion();
}

/**
 * Move to the next suggestion
 */
function moveToNextSuggestion() {
    currentSuggestionIndex++;
    setButtonsEnabled(true);
    displayCurrentSuggestion();
}

/**
 * Show completion state when all suggestions are processed
 */
function showCompletionState() {
    document.getElementById("mainContent").classList.add("hidden");
    document.getElementById("progressContainer").classList.add("hidden");
    document.getElementById("emptyState").classList.remove("hidden");
    
    const appliedCount = processedSuggestions.filter(p => p.applied).length;
    const skippedCount = processedSuggestions.filter(p => p.skipped).length;
    
    showStatus(`Completed! Applied: ${appliedCount}, Skipped: ${skippedCount}`, "success");
}

/**
 * Enable/disable action buttons
 */
function setButtonsEnabled(enabled) {
    document.getElementById("applyButton").disabled = !enabled;
    document.getElementById("skipButton").disabled = !enabled;
    
    if (!enabled) {
        document.getElementById("applyButton").innerHTML = 
            '<span class="loading-spinner"></span>Applying...';
    } else {
        document.getElementById("applyButton").innerHTML = 'Apply Suggestion';
    }
}

/**
 * Show status message
 */
function showStatus(message, type = "info") {
    const statusElement = document.getElementById("statusMessage");
    statusElement.textContent = message;
    statusElement.className = `status ${type}`;
    statusElement.classList.remove("hidden");
    
    // Auto-hide after 5 seconds for non-error messages
    if (type !== "error") {
        setTimeout(() => {
            statusElement.classList.add("hidden");
        }, 5000);
    }
}

// Utility functions for debugging
window.debugAddin = {
    getSuggestions: () => suggestions,
    getCurrentIndex: () => currentSuggestionIndex,
    getProcessedSuggestions: () => processedSuggestions,
    resetAddin: () => {
        suggestions = null;
        currentSuggestionIndex = 0;
        processedSuggestions = [];
        document.getElementById("loadSection").classList.remove("hidden");
        document.getElementById("mainContent").classList.add("hidden");
        document.getElementById("progressContainer").classList.add("hidden");
        document.getElementById("emptyState").classList.add("hidden");
        showStatus("Add-in reset. Load suggestions to start again.", "info");
    }
}; 