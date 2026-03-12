# Across the Office - Task Notes

## Current Status
**Windows-only release.** Signing via Azure Artifact Signing (formerly Trusted Signing).

## Azure Artifact Signing Setup

The CI workflow signs the Windows exe using Azure Artifact Signing. Three GitHub secrets are required:
- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET` (App Registration client secret, **expires every 90 days**)

### App Registration Role Assignment
The App Registration's service principal must have the **"Artifact Signing Certificate Profile Signer"** role on the Trusted Signing Account (`orfe-signing-acct`):
1. Azure portal → Trusted Signing Account → Access control (IAM)
2. Add role assignment → **Artifact Signing Certificate Profile Signer**
3. Assign to the App Registration's service principal

### Rotating the Client Secret (every 90 days)
1. Azure portal → App Registrations → your app → Certificates & secrets
2. Create a new client secret
3. Update the GitHub secret: `gh secret set AZURE_CLIENT_SECRET`
4. Delete the old secret in Azure

## CI/CD Status
- Build and Release workflow: Windows portable exe only (macOS disabled)
- GitHub Pages workflow: Working (deploys docs/ folder)
- Signing: Automatic when Azure secrets are present, unsigned otherwise

## To Create a Release
```bash
git tag -a vX.Y.Z -m "Release vX.Y.Z"
git push origin vX.Y.Z
```

## Run Commands
```bash
npm start         # Run the app
npm run dist:win  # Build for Windows
```
