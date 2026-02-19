const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');

let mainWindow;
let reportWindow = null;
let reportData = null;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 900,
    height: 700,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  mainWindow.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// Handle folder selection
ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory']
  });
  return result.canceled ? null : result.filePaths[0];
});

// Reveal file in system file browser (Finder, Explorer, etc.)
ipcMain.handle('reveal-file', async (event, filePath) => {
  shell.showItemInFolder(filePath);
});

// Open report window with dry run results
ipcMain.handle('open-report-window', async (event, data) => {
  reportData = data;

  if (reportWindow && !reportWindow.isDestroyed()) {
    reportWindow.focus();
    reportWindow.webContents.reload();
    return;
  }

  reportWindow = new BrowserWindow({
    width: 800,
    height: 600,
    parent: mainWindow,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'report-preload.js')
    }
  });

  reportWindow.loadFile('report.html');

  reportWindow.on('closed', () => {
    reportWindow = null;
  });
});

// Provide report data to the report window
ipcMain.handle('get-report-data', async () => {
  return reportData;
});

// Extract text snippets from a Word document for preview
async function extractWordSnippets(filePath, searchText, matchCase) {
  try {
    const content = fs.readFileSync(filePath);
    const zip = new PizZip(content);
    const documentXml = zip.file('word/document.xml');

    if (!documentXml) return [];

    const xmlContent = documentXml.asText();
    const snippets = [];

    // Extract full text
    const textRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let match;
    let fullText = '';

    while ((match = textRegex.exec(xmlContent)) !== null) {
      fullText += match[1];
    }

    // Find matches and extract context
    const regexFlags = matchCase ? 'g' : 'gi';
    const regex = new RegExp(escapeRegex(searchText), regexFlags);
    const contextChars = 40;

    while ((match = regex.exec(fullText)) !== null && snippets.length < 5) {
      const start = Math.max(0, match.index - contextChars);
      const end = Math.min(fullText.length, match.index + match[0].length + contextChars);

      const before = (start > 0 ? '...' : '') + fullText.substring(start, end) + (end < fullText.length ? '...' : '');
      const after = before.replace(new RegExp(escapeRegex(searchText), regexFlags), (m) => m);

      snippets.push({ before, after });
    }

    return snippets;
  } catch (error) {
    console.error(`Error extracting snippets from ${filePath}:`, error);
    return [];
  }
}

// Extract text snippets from a PowerPoint document for preview
async function extractPptSnippets(filePath, searchText, matchCase) {
  try {
    const content = fs.readFileSync(filePath);
    const zip = new PizZip(content);
    const snippets = [];

    const slideFiles = Object.keys(zip.files).filter(name =>
      name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
    );

    const regexFlags = matchCase ? 'g' : 'gi';
    const contextChars = 40;

    for (const slideFile of slideFiles) {
      if (snippets.length >= 5) break;

      const slideContent = zip.file(slideFile);
      if (!slideContent) continue;

      const xmlContent = slideContent.asText();
      const textRegex = /<a:t>([^<]*)<\/a:t>/g;
      let match;
      let slideText = '';

      while ((match = textRegex.exec(xmlContent)) !== null) {
        slideText += match[1];
      }

      const regex = new RegExp(escapeRegex(searchText), regexFlags);

      while ((match = regex.exec(slideText)) !== null && snippets.length < 5) {
        const start = Math.max(0, match.index - contextChars);
        const end = Math.min(slideText.length, match.index + match[0].length + contextChars);

        const before = (start > 0 ? '...' : '') + slideText.substring(start, end) + (end < slideText.length ? '...' : '');
        const after = before.replace(new RegExp(escapeRegex(searchText), regexFlags), (m) => m);

        snippets.push({ before, after });
      }
    }

    return snippets;
  } catch (error) {
    console.error(`Error extracting snippets from ${filePath}:`, error);
    return [];
  }
}

// Recursively find all Word and PowerPoint files with progress callback
function findOfficeFiles(dir, files = [], onProgress = null, dirCount = { count: 0 }) {
  try {
    const entries = fs.readdirSync(dir, { withFileTypes: true });
    dirCount.count++;

    // Send progress update for directory scanning
    if (onProgress && dirCount.count % 10 === 0) {
      onProgress({ phase: 'scanning', dirsScanned: dirCount.count, filesFound: files.length });
    }

    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);

      if (entry.isDirectory()) {
        // Skip hidden directories and common system directories
        if (!entry.name.startsWith('.') && entry.name !== 'node_modules') {
          findOfficeFiles(fullPath, files, onProgress, dirCount);
        }
      } else {
        const ext = path.extname(entry.name).toLowerCase();
        if (['.docx', '.pptx'].includes(ext)) {
          files.push(fullPath);
          if (onProgress) {
            onProgress({ phase: 'scanning', dirsScanned: dirCount.count, filesFound: files.length });
          }
        }
      }
    }
  } catch (error) {
    // Skip directories we can't access
    console.error(`Error accessing directory ${dir}:`, error.message);
  }

  return files;
}

// Search for text in Word document (.docx)
async function searchWordDocument(filePath, searchText, matchCase) {
  try {
    const content = fs.readFileSync(filePath);
    const zip = new PizZip(content);
    const documentXml = zip.file('word/document.xml');

    if (!documentXml) return { found: false, matches: [] };

    const xmlContent = documentXml.asText();
    const matches = [];

    // Extract text content and search
    const textRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let match;
    let fullText = '';

    while ((match = textRegex.exec(xmlContent)) !== null) {
      fullText += match[1];
    }

    const textToSearch = matchCase ? fullText : fullText.toLowerCase();
    const textToFind = matchCase ? searchText : searchText.toLowerCase();

    if (textToSearch.includes(textToFind)) {
      // Count occurrences
      const regexFlags = matchCase ? 'g' : 'gi';
      const regex = new RegExp(escapeRegex(searchText), regexFlags);
      const count = (fullText.match(regex) || []).length;
      matches.push({ count });
    }

    return { found: matches.length > 0, matches, fullText };
  } catch (error) {
    console.error(`Error reading Word document ${filePath}:`, error);
    return { found: false, error: error.message };
  }
}

// Search for text in PowerPoint document (.pptx)
async function searchPowerPointDocument(filePath, searchText, matchCase) {
  try {
    const content = fs.readFileSync(filePath);
    const zip = new PizZip(content);
    const matches = [];
    let totalCount = 0;

    // Search through all slides
    const slideFiles = Object.keys(zip.files).filter(name =>
      name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
    );

    const regexFlags = matchCase ? 'g' : 'gi';

    for (const slideFile of slideFiles) {
      const slideContent = zip.file(slideFile);
      if (!slideContent) continue;

      const xmlContent = slideContent.asText();
      const textRegex = /<a:t>([^<]*)<\/a:t>/g;
      let match;
      let slideText = '';

      while ((match = textRegex.exec(xmlContent)) !== null) {
        slideText += match[1];
      }

      const textToSearch = matchCase ? slideText : slideText.toLowerCase();
      const textToFind = matchCase ? searchText : searchText.toLowerCase();

      if (textToSearch.includes(textToFind)) {
        const regex = new RegExp(escapeRegex(searchText), regexFlags);
        const count = (slideText.match(regex) || []).length;
        if (count > 0) {
          const slideNum = slideFile.match(/slide(\d+)\.xml/)[1];
          matches.push({ slide: parseInt(slideNum), count });
          totalCount += count;
        }
      }
    }

    return { found: matches.length > 0, matches, totalCount };
  } catch (error) {
    console.error(`Error reading PowerPoint document ${filePath}:`, error);
    return { found: false, error: error.message };
  }
}

function escapeRegex(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Escape XML special characters to prevent breaking document structure
function escapeXml(string) {
  return string
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Handle search operation
ipcMain.handle('search-files', async (event, { folderPath, searchText, matchCase }) => {
  const results = [];
  const sender = event.sender;

  // Progress callback for directory scanning
  const onScanProgress = (progress) => {
    sender.send('search-progress', progress);
  };

  try {
    // Phase 1: Scan directories for Office files
    sender.send('search-progress', { phase: 'scanning', dirsScanned: 0, filesFound: 0 });
    const files = findOfficeFiles(folderPath, [], onScanProgress);

    // Phase 2: Search through each file
    const totalFiles = files.length;
    sender.send('search-progress', { phase: 'searching', current: 0, total: totalFiles, matchesFound: 0 });

    for (let i = 0; i < files.length; i++) {
      const filePath = files[i];
      const ext = path.extname(filePath).toLowerCase();
      let searchResult;

      // Send progress update
      sender.send('search-progress', {
        phase: 'searching',
        current: i + 1,
        total: totalFiles,
        currentFile: path.basename(filePath),
        matchesFound: results.length
      });

      if (ext === '.docx') {
        searchResult = await searchWordDocument(filePath, searchText, matchCase);
      } else if (ext === '.pptx') {
        searchResult = await searchPowerPointDocument(filePath, searchText, matchCase);
      }

      if (searchResult && searchResult.found) {
        results.push({
          filePath,
          fileName: path.basename(filePath),
          fileType: ext === '.docx' ? 'Word' : 'PowerPoint',
          matches: searchResult.matches,
          totalCount: searchResult.totalCount || searchResult.matches[0]?.count || 0
        });
      }
    }

    // Signal completion
    sender.send('search-progress', { phase: 'complete' });
  } catch (error) {
    console.error('Search error:', error);
    throw error;
  }

  return results;
});

// Replace text in Word document - handles text spanning multiple w:t elements
async function replaceInWordDocument(filePath, searchText, replaceText, matchCase) {
  const content = fs.readFileSync(filePath);
  const zip = new PizZip(content);
  const documentXml = zip.file('word/document.xml');

  if (!documentXml) return false;

  let xmlContent = documentXml.asText();

  // Strategy: process each paragraph (w:p) element separately to handle text spans
  // within the same paragraph while preserving paragraph structure
  xmlContent = xmlContent.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, (paragraph) => {
    return replaceTextInWordParagraph(paragraph, searchText, replaceText, matchCase);
  });

  zip.file('word/document.xml', xmlContent);
  return zip.generate({ type: 'nodebuffer' });
}

// Replace text within a Word paragraph, preserving formatting on non-matched text
function replaceTextInWordParagraph(paragraph, searchText, replaceText, matchCase) {
  const regexFlags = matchCase ? 'g' : 'gi';
  const safeReplaceText = escapeXml(replaceText);

  // Extract all w:t elements with character position tracking
  const elements = [];
  const tagRegex = /<w:t([^>]*)>([^<]*)<\/w:t>/g;
  let match;
  let charPos = 0;

  while ((match = tagRegex.exec(paragraph)) !== null) {
    elements.push({
      fullMatch: match[0],
      attrs: match[1],
      text: match[2],
      charStart: charPos,
      charEnd: charPos + match[2].length
    });
    charPos += match[2].length;
  }

  if (elements.length === 0) return paragraph;

  const fullText = elements.map(e => e.text).join('');
  const searchRegex = new RegExp(escapeRegex(searchText), regexFlags);

  // Find all matches
  const matches = [];
  while ((match = searchRegex.exec(fullText)) !== null) {
    matches.push({ start: match.index, end: match.index + match[0].length });
  }

  if (matches.length === 0) return paragraph;

  // Build action array: 0=keep char, 1=delete char, 2=insert replacement here
  const actions = new Array(fullText.length).fill(0);
  for (const m of matches) {
    actions[m.start] = 2; // First char of match triggers replacement
    for (let i = m.start + 1; i < m.end; i++) {
      actions[i] = 1; // Rest of matched chars are deleted
    }
  }

  // Generate new text for each element based on actions
  let result = paragraph;
  for (const elem of elements) {
    let newText = '';
    for (let i = elem.charStart; i < elem.charEnd; i++) {
      if (actions[i] === 0) {
        newText += fullText[i]; // Keep original char
      } else if (actions[i] === 2) {
        newText += safeReplaceText; // Insert replacement
      }
      // actions[i] === 1: skip (delete char)
    }
    result = result.replace(elem.fullMatch, `<w:t${elem.attrs}>${newText}</w:t>`);
  }

  return result;
}

// Replace text in PowerPoint document - handles text spanning multiple a:t elements
async function replaceInPowerPointDocument(filePath, searchText, replaceText, matchCase) {
  const content = fs.readFileSync(filePath);
  const zip = new PizZip(content);

  const slideFiles = Object.keys(zip.files).filter(name =>
    name.startsWith('ppt/slides/slide') && name.endsWith('.xml')
  );

  for (const slideFile of slideFiles) {
    const slideContent = zip.file(slideFile);
    if (!slideContent) continue;

    let xmlContent = slideContent.asText();

    // Process each text paragraph (a:p) element to handle text spans
    xmlContent = xmlContent.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g, (paragraph) => {
      return replaceTextInPptParagraph(paragraph, searchText, replaceText, matchCase);
    });

    zip.file(slideFile, xmlContent);
  }

  return zip.generate({ type: 'nodebuffer' });
}

// Replace text within a PowerPoint paragraph, preserving formatting on non-matched text
function replaceTextInPptParagraph(paragraph, searchText, replaceText, matchCase) {
  const regexFlags = matchCase ? 'g' : 'gi';
  const safeReplaceText = escapeXml(replaceText);

  // Extract all a:t elements with character position tracking
  const elements = [];
  const tagRegex = /<a:t>([^<]*)<\/a:t>/g;
  let match;
  let charPos = 0;

  while ((match = tagRegex.exec(paragraph)) !== null) {
    elements.push({
      fullMatch: match[0],
      text: match[1],
      charStart: charPos,
      charEnd: charPos + match[1].length
    });
    charPos += match[1].length;
  }

  if (elements.length === 0) return paragraph;

  const fullText = elements.map(e => e.text).join('');
  const searchRegex = new RegExp(escapeRegex(searchText), regexFlags);

  // Find all matches
  const matches = [];
  while ((match = searchRegex.exec(fullText)) !== null) {
    matches.push({ start: match.index, end: match.index + match[0].length });
  }

  if (matches.length === 0) return paragraph;

  // Build action array: 0=keep char, 1=delete char, 2=insert replacement here
  const actions = new Array(fullText.length).fill(0);
  for (const m of matches) {
    actions[m.start] = 2; // First char of match triggers replacement
    for (let i = m.start + 1; i < m.end; i++) {
      actions[i] = 1; // Rest of matched chars are deleted
    }
  }

  // Generate new text for each element based on actions
  let result = paragraph;
  for (const elem of elements) {
    let newText = '';
    for (let i = elem.charStart; i < elem.charEnd; i++) {
      if (actions[i] === 0) {
        newText += fullText[i]; // Keep original char
      } else if (actions[i] === 2) {
        newText += safeReplaceText; // Insert replacement
      }
      // actions[i] === 1: skip (delete char)
    }
    result = result.replace(elem.fullMatch, `<a:t>${newText}</a:t>`);
  }

  return result;
}

// Generate modified filename with identifier
function generateModifiedPath(filePath) {
  const dir = path.dirname(filePath);
  const ext = path.extname(filePath);
  const baseName = path.basename(filePath, ext);
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  return path.join(dir, `${baseName}_modified_by_ato_${timestamp}${ext}`);
}

// Handle replace operation
ipcMain.handle('replace-in-files', async (event, { files, searchText, replaceText, matchCase, dryRun }) => {
  const results = [];

  for (const file of files) {
    try {
      const ext = path.extname(file.filePath).toLowerCase();
      let newContent;

      if (ext === '.docx') {
        newContent = await replaceInWordDocument(file.filePath, searchText, replaceText, matchCase);
      } else if (ext === '.pptx') {
        newContent = await replaceInPowerPointDocument(file.filePath, searchText, replaceText, matchCase);
      }

      if (newContent) {
        const modifiedPath = generateModifiedPath(file.filePath);

        if (dryRun) {
          // Dry run: report what would happen without writing
          // Extract snippets for preview
          let snippets = [];
          if (ext === '.docx') {
            snippets = await extractWordSnippets(file.filePath, searchText, matchCase);
          } else if (ext === '.pptx') {
            snippets = await extractPptSnippets(file.filePath, searchText, matchCase);
          }

          // Generate before/after for each snippet
          const regexFlags = matchCase ? 'g' : 'gi';
          const snippetsWithAfter = snippets.map(s => ({
            before: s.before,
            after: s.before.replace(new RegExp(escapeRegex(searchText), regexFlags), replaceText)
          }));

          results.push({
            filePath: file.filePath,
            fileName: file.fileName,
            fileType: file.fileType,
            success: true,
            dryRun: true,
            modifiedPath,
            replacementCount: file.totalCount,
            snippets: snippetsWithAfter
          });
        } else {
          // Write modified content to new file, leaving original untouched
          fs.writeFileSync(modifiedPath, newContent);

          results.push({
            filePath: file.filePath,
            success: true,
            modifiedPath
          });
        }
      }
    } catch (error) {
      console.error(`Error replacing in ${file.filePath}:`, error);
      results.push({
        filePath: file.filePath,
        success: false,
        error: error.message
      });
    }
  }

  return results;
});
