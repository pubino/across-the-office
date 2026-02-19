# Across the Office - Task Notes

## Current Status
**All requirements implemented.** Ready for manual verification and testing.

## Manual Testing Checklist
Before marking project complete, verify with real Office files:
- [ ] Search finds text in Word (.docx) documents
- [ ] Search finds text in PowerPoint (.pptx) documents
- [ ] Replace creates `_modified_by_ato_` files, leaving originals untouched
- [ ] Replace works with special chars like `Tom & Jerry` or `<Company>`
- [ ] Select All / None buttons work correctly
- [ ] Individual file selection works
- [ ] Match case option works correctly
- [ ] Dry run shows preview report without modifying files

## Run Commands
```bash
npm start     # Run the app
```
