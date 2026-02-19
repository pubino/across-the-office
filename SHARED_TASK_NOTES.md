# Across the Office - Task Notes

## Current Status
**Project complete.** All primary goals achieved.

## Completed
- Security audit (passed - good Electron security practices)
- README.md with app description and unsigned release instructions
- LICENSE.md (MIT, Princeton University)
- CI/CD pipeline for macOS and Windows builds (.github/workflows/build.yml)
- GitHub Pages site (docs/index.html, .github/workflows/pages.yml)
- Git repo initialized and pushed to https://github.com/pubino/across-the-office
- GitHub Pages published at https://pubino.github.io/across-the-office/

## CI/CD Status
- Build and Release workflow: Working (builds macOS dmg/zip and Windows exe)
- GitHub Pages workflow: Working (deploys docs/ folder)

## To Create a Release
```bash
git tag v1.0.0
git push origin v1.0.0
```
This will trigger the release job which creates a GitHub Release with all build artifacts.

## Run Commands
```bash
npm start     # Run the app
npm run dist  # Build for current platform
```
