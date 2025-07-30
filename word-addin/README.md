# Style Suggestions Word Add-in

A Microsoft Word add-in that applies AI-powered style suggestions to improve document formatting based on template comparisons.

## Features

- **Interactive Style Suggestions**: Load and review style suggestions with visual previews
- **Document Navigation**: Automatically navigate to paragraphs that need style changes
- **Live Preview**: See current vs. suggested formatting before applying changes
- **Progress Tracking**: Track progress through all suggestions with visual progress bar
- **Batch Processing**: Apply or skip suggestions one by one with full control

## Getting Started

### Prerequisites

- Node.js (version 14 or higher)
- Microsoft Word (Office 365 or Word 2016+)
- Office Add-ins Development Tools

### Installation

1. Clone or download this add-in to your local machine
2. Navigate to the `word-addin` directory
3. Install dependencies:
   ```bash
   npm install
   ```

### Development

1. Start the development server:
   ```bash
   npm run dev-server
   ```

2. In another terminal, start the add-in:
   ```bash
   npm start
   ```

3. Word will open with the add-in loaded in the task pane

### Building for Production

```bash
npm run build
```

## Usage

### Loading Suggestions

1. Open the add-in in Word
2. Click "Load Style Suggestions" 
3. The add-in will load demo suggestions (or you can modify to load from a file/API)

### Reviewing Suggestions

For each suggestion, you'll see:
- **Current Style**: How the text currently appears
- **Suggested Style**: How it would look after applying the suggestion
- **Context**: Location information (paragraph number, section, etc.)
- **Explanation**: Why this change is recommended

### Applying Changes

- **Apply Suggestion**: Applies the suggested formatting and moves to the next suggestion
- **Skip**: Skips the current suggestion and moves to the next one
- **Auto-Navigation**: The add-in automatically navigates to each paragraph in Word

## JSON Format

The add-in expects suggestions in this format:

```json
{
  "suggestions": {
    "document": [
      {
        "json_object": [
          {
            "id": 277,
            "formattingContext": {
              "sampleText": "Sample text content",
              "paragraphIndex": 36,
              "sectionIndex": 0,
              "contextKey": "Section:0:Paragraph:36"
            },
            "fontFamily": "Current Font",
            "color": "4F81BD",
            "basedOnStyle": "Heading4"
          }
        ],
        "ops": [
          {
            "prop": "paragraph.style",
            "to": "Heading2"
          }
        ],
        "message": "Explanation of why this change is suggested"
      }
    ]
  }
}
```

### Supported Operations

- `paragraph.style`: Change paragraph style
- `font.name`: Change font family
- `style.font.color`: Change font color
- `paragraph.alignment`: Change text alignment

## Architecture

### Files Structure

```
word-addin/
├── manifest.xml              # Add-in manifest
├── package.json              # Dependencies and scripts
├── webpack.config.js         # Build configuration
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html     # Main UI
│   │   └── taskpane.js       # Main logic
│   └── commands/
│       ├── commands.html     # Command functions
│       └── commands.js       # Command handlers
├── assets/                   # Icons and images
└── dist/                     # Built files (generated)
```

### Key Components

1. **SuggestionProcessor**: Handles loading and processing of JSON suggestions
2. **WordNavigator**: Manages navigation to specific paragraphs in Word
3. **StyleApplier**: Applies formatting changes using Word JavaScript API
4. **UIController**: Manages the task pane interface and user interactions

## Development Features

### Debug Console

The add-in includes debug utilities accessible via browser console:

```javascript
// Get current suggestions
debugAddin.getSuggestions()

// Get current index
debugAddin.getCurrentIndex()

// Get processed suggestions history
debugAddin.getProcessedSuggestions()

// Reset the add-in
debugAddin.resetAddin()
```

### Error Handling

- Graceful handling of missing paragraphs
- User-friendly error messages
- Automatic recovery from Word API errors
- Validation of JSON structure

## Integration with Style Verification Service

This add-in is designed to work with the Style Verification Service API:

1. **Document Analysis**: The service analyzes document styles
2. **Template Comparison**: Compares against template styles
3. **Suggestion Generation**: Creates JSON with style suggestions
4. **Add-in Processing**: This add-in applies the suggestions

### Future Enhancements

- **API Integration**: Direct connection to style verification service
- **Batch Operations**: Apply all suggestions at once
- **Undo Functionality**: Revert applied changes
- **Custom Templates**: Load different template styles
- **Export Reports**: Generate reports of applied changes

## Troubleshooting

### Common Issues

1. **Add-in not loading**: Ensure dev server is running on port 3000
2. **Paragraph not found**: Check that paragraph indices match document structure
3. **Style not applied**: Verify style names exist in the document
4. **Navigation issues**: Ensure Word document is active and accessible

### Debugging

1. Open browser developer tools (F12 in Word)
2. Check console for error messages
3. Use debug utilities to inspect add-in state
4. Verify JSON structure matches expected format

## License

MIT License - see LICENSE file for details

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review browser console for errors
3. Verify JSON format compliance
4. Test with demo suggestions first

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes and test thoroughly
4. Submit a pull request with detailed description

---

**Note**: This add-in is designed to work with the Style Verification Service and expects specific JSON formats. Ensure your suggestion data matches the expected structure for optimal performance. 