# Across the Office

A desktop application for finding and replacing text across Microsoft Office documents (Word `.docx` and PowerPoint `.pptx` files).

## Features

- **Batch Search**: Recursively scan folders for Office documents containing specific text
- **Safe Replace**: Create modified copies while preserving original files
- **Dry Run Mode**: Preview changes before applying them
- **Case Sensitivity**: Optional case-sensitive matching
- **Cross-Platform**: Works on macOS and Windows

## Installation

### Download Pre-built Releases

Download the latest release for your platform from the [Releases page](https://github.com/pubino/across-the-office/releases).

### Build from Source

```bash
# Clone the repository
git clone https://github.com/pubino/across-the-office.git
cd across-the-office

# Install dependencies
npm install

# Run the application
npm start

# Build for your platform
npm run dist
```

## Usage

1. **Select a folder** containing Office documents
2. **Enter search text** to find in documents
3. **Click Search** to scan all documents
4. **Review results** showing files with matches
5. **Select files** to modify (or use Select All/None)
6. **Enter replacement text** (optional)
7. **Enable Dry Run** to preview changes (recommended first)
8. **Click Replace** to create modified copies

Modified files are saved with `_modified_by_ato_TIMESTAMP` suffix, leaving originals untouched.

## Unsigned Releases

The pre-built releases are **not code-signed**. On macOS, you may see a warning that the app is from an unidentified developer. To open it:

1. Right-click (or Control-click) the app
2. Select "Open" from the context menu
3. Click "Open" in the dialog that appears

On Windows, you may see a SmartScreen warning. Click "More info" and then "Run anyway".

### Building Signed Releases

If you require signed applications (for enterprise deployment or to avoid security warnings), you can build and sign the app yourself:

**macOS:**
```bash
# Set your signing identity
export CSC_NAME="Developer ID Application: Your Name (TEAM_ID)"

# Build signed app
npm run dist
```

You'll need an Apple Developer account and valid Developer ID certificate.

**Windows:**
```bash
# Set your certificate path and password
export CSC_LINK="path/to/certificate.pfx"
export CSC_KEY_PASSWORD="your-password"

# Build signed app
npm run dist
```

You'll need a code signing certificate from a trusted Certificate Authority.

## Development

```bash
# Install dependencies
npm install

# Run in development mode
npm start

# Run tests
npm test

# Build distributable packages
npm run dist
```

## How It Works

Office documents (`.docx` and `.pptx`) are ZIP archives containing XML files. This application:

1. Extracts the XML content from each document
2. Searches for text within the XML structure
3. Performs replacements while preserving document formatting
4. Saves the modified XML back into a new document file

## Security

- Original files are never modified
- All user input is properly escaped before processing
- The application runs with minimal permissions
- No network access required

## License

[MIT License](LICENSE.md) - Copyright (c) 2026 Princeton University

## Contributing

Contributions are welcome! Please feel free to submit issues and pull requests.
