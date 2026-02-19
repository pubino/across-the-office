# Across the Office

A desktop application for finding and replacing text across Microsoft Office documents (Word `.docx` and PowerPoint `.pptx` files).

## Features

- **Batch Search**: Recursively scan folders for Office documents containing specific text
- **Safe Replace**: Create modified copies while preserving original files
- **Dry Run Mode**: Preview changes before applying them
- **Case Sensitivity**: Optional case-sensitive matching
- **Windows App**: Portable executable—no installation required

## Installation

### Download Pre-built Release

Download the latest Windows release from the [Releases page](https://github.com/pubino/across-the-office/releases).

### Build from Source

```bash
# Clone the repository
git clone https://github.com/pubino/across-the-office.git
cd across-the-office

# Install dependencies
npm install

# Run the application
npm start

# Build for Windows
npm run dist:win
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

## Code Signing

The pre-built releases are **not code-signed**. If your organization requires signed applications, you can build and sign the app yourself:

1. Obtain a code signing certificate from a trusted Certificate Authority (e.g., DigiCert, Sectigo, or SSL.com)

2. Set your certificate environment variables:
```cmd
set CSC_LINK=path\to\certificate.pfx
set CSC_KEY_PASSWORD=your-password
```

3. Build the signed application:
```cmd
npm run dist:win
```

The signed executable will be created in the `dist` folder.

## Development

```bash
# Install dependencies
npm install

# Run in development mode
npm start

# Run tests
npm test

# Build Windows executable
npm run dist:win
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

