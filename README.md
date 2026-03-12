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

Releases are signed via [Azure Artifact Signing](https://azure.microsoft.com/en-us/products/artifact-signing) (formerly Trusted Signing). The CI workflow signs automatically when the required GitHub secrets are configured.

### CI Signing Setup

1. Create an Azure Artifact Signing account and certificate profile
2. Register an App in Azure Entra ID (App Registrations)
3. Assign the **"Artifact Signing Certificate Profile Signer"** role to the App Registration on the signing account (Access control → Add role assignment)
4. Create a client secret for the App Registration (note: secrets expire — see below)
5. Add these GitHub repository secrets (Settings → Secrets → Actions):
   - `AZURE_TENANT_ID` — from Azure Entra ID overview
   - `AZURE_CLIENT_ID` — from the App Registration overview
   - `AZURE_CLIENT_SECRET` — the client secret value

### Rotating the Client Secret

The Azure App Registration client secret has an expiry (e.g., 90 days). To rotate:

1. Azure portal → App Registrations → your app → Certificates & secrets
2. Create a new client secret
3. Update the GitHub secret: `gh secret set AZURE_CLIENT_SECRET`
4. Delete the old secret in Azure

### Local Signing

To build a signed executable locally, set the Azure environment variables and run:

```cmd
set AZURE_TENANT_ID=your-tenant-id
set AZURE_CLIENT_ID=your-client-id
set AZURE_CLIENT_SECRET=your-client-secret
npm run dist:win
```

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

