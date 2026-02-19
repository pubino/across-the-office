// DOM elements
const searchTextInput = document.getElementById('searchText');
const replaceTextInput = document.getElementById('replaceText');
const matchCaseCheckbox = document.getElementById('matchCase');
const dryRunCheckbox = document.getElementById('dryRun');
const folderPathInput = document.getElementById('folderPath');
const selectFolderBtn = document.getElementById('selectFolderBtn');
const searchBtn = document.getElementById('searchBtn');
const resultsSection = document.getElementById('resultsSection');
const fileList = document.getElementById('fileList');
const resultCount = document.getElementById('resultCount');
const selectAllBtn = document.getElementById('selectAllBtn');
const selectNoneBtn = document.getElementById('selectNoneBtn');
const replaceBtn = document.getElementById('replaceBtn');
const statusMessage = document.getElementById('statusMessage');
const searchLoading = document.getElementById('searchLoading');
const replaceLoading = document.getElementById('replaceLoading');
const progressPhase = document.getElementById('progressPhase');
const progressBar = document.getElementById('progressBar');
const progressDetails = document.getElementById('progressDetails');

// State
let searchResults = [];
let selectedFolder = null;

// Utility functions
function showStatus(message, type = 'info') {
  statusMessage.textContent = message;
  statusMessage.className = `status-message visible ${type}`;
}

function hideStatus() {
  statusMessage.className = 'status-message';
}

function updateSearchButton() {
  const hasSearchText = searchTextInput.value.trim().length > 0;
  const hasFolder = selectedFolder !== null;
  searchBtn.disabled = !(hasSearchText && hasFolder);
}

function updateReplaceButton() {
  const checkedBoxes = fileList.querySelectorAll('input[type="checkbox"]:checked');
  replaceBtn.disabled = checkedBoxes.length === 0;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function updateProgress(data) {
  if (data.phase === 'scanning') {
    progressPhase.textContent = 'Scanning folders...';
    progressBar.style.width = '0%';
    progressBar.classList.add('indeterminate');
    progressDetails.textContent = `${data.dirsScanned} folders scanned, ${data.filesFound} Office files found`;
  } else if (data.phase === 'searching') {
    progressPhase.textContent = 'Searching files...';
    progressBar.classList.remove('indeterminate');
    const percent = data.total > 0 ? Math.round((data.current / data.total) * 100) : 0;
    progressBar.style.width = `${percent}%`;
    let details = `${data.current} of ${data.total} files`;
    if (data.currentFile) {
      details += ` - ${data.currentFile}`;
    }
    if (data.matchesFound > 0) {
      details += ` (${data.matchesFound} matches found)`;
    }
    progressDetails.textContent = details;
  } else if (data.phase === 'complete') {
    progressBar.style.width = '100%';
  }
}

// Event handlers
selectFolderBtn.addEventListener('click', async () => {
  const folder = await window.electronAPI.selectFolder();
  if (folder) {
    selectedFolder = folder;
    folderPathInput.value = folder;
    updateSearchButton();
    hideStatus();
  }
});

searchTextInput.addEventListener('input', updateSearchButton);

searchBtn.addEventListener('click', async () => {
  const searchText = searchTextInput.value.trim();
  if (!searchText || !selectedFolder) return;

  hideStatus();
  searchLoading.classList.add('visible');
  searchBtn.disabled = true;
  resultsSection.classList.remove('visible');

  // Reset progress UI
  progressBar.style.width = '0%';
  progressBar.classList.add('indeterminate');
  progressPhase.textContent = 'Scanning folders...';
  progressDetails.textContent = '';

  // Set up progress listener
  window.electronAPI.onSearchProgress(updateProgress);

  try {
    const matchCase = matchCaseCheckbox.checked;
    searchResults = await window.electronAPI.searchFiles(selectedFolder, searchText, matchCase);

    if (searchResults.length === 0) {
      showStatus('No files found containing the search text.', 'info');
    } else {
      renderResults();
      resultsSection.classList.add('visible');
    }
  } catch (error) {
    showStatus(`Error during search: ${error.message}`, 'error');
  } finally {
    // Clean up progress listener
    window.electronAPI.removeSearchProgressListener();
    searchLoading.classList.remove('visible');
    updateSearchButton();
  }
});

function renderResults() {
  fileList.innerHTML = '';
  resultCount.textContent = searchResults.length;

  searchResults.forEach((result, index) => {
    const item = document.createElement('div');
    item.className = 'file-item';

    const typeClass = result.fileType === 'Word' ? 'word' : 'powerpoint';
    const matchText = result.totalCount === 1 ? '1 match' : `${result.totalCount} matches`;

    item.innerHTML = `
      <input type="checkbox" id="file-${index}" data-index="${index}" checked>
      <div class="file-info">
        <div class="file-name">${escapeHtml(result.fileName)}</div>
        <div class="file-path clickable" data-path="${escapeHtml(result.filePath)}">${escapeHtml(result.filePath)}</div>
      </div>
      <span class="file-type ${typeClass}">${result.fileType}</span>
      <span class="match-count">${matchText}</span>
    `;

    const checkbox = item.querySelector('input[type="checkbox"]');
    checkbox.addEventListener('change', updateReplaceButton);

    const pathElement = item.querySelector('.file-path');
    pathElement.addEventListener('click', () => {
      window.electronAPI.revealFile(result.filePath);
    });

    fileList.appendChild(item);
  });

  updateReplaceButton();
}

selectAllBtn.addEventListener('click', () => {
  const checkboxes = fileList.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(cb => cb.checked = true);
  updateReplaceButton();
});

selectNoneBtn.addEventListener('click', () => {
  const checkboxes = fileList.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(cb => cb.checked = false);
  updateReplaceButton();
});

replaceBtn.addEventListener('click', async () => {
  const searchText = searchTextInput.value.trim();
  const replaceText = replaceTextInput.value;
  const checkedBoxes = fileList.querySelectorAll('input[type="checkbox"]:checked');

  if (checkedBoxes.length === 0) {
    showStatus('No files selected for replacement.', 'error');
    return;
  }

  const selectedFiles = Array.from(checkedBoxes).map(cb => {
    const index = parseInt(cb.dataset.index);
    return searchResults[index];
  });

  hideStatus();
  replaceLoading.classList.add('visible');
  replaceBtn.disabled = true;

  try {
    const matchCase = matchCaseCheckbox.checked;
    const dryRun = dryRunCheckbox.checked;
    const results = await window.electronAPI.replaceInFiles(selectedFiles, searchText, replaceText, matchCase, dryRun);

    const successful = results.filter(r => r.success);
    const failed = results.filter(r => !r.success);

    if (dryRun) {
      // Open dry run report in new window
      await window.electronAPI.openReportWindow({
        searchText,
        replaceText,
        results: successful
      });
      showStatus('Dry run complete. Report opened in new window.', 'info');
    } else {
      // Actual replacement
      let message = `Created ${successful.length} modified file(s).`;
      if (failed.length > 0) {
        message += ` ${failed.length} file(s) failed.`;
      }
      message += ' Originals left untouched.';

      showStatus(message, failed.length > 0 ? 'error' : 'success');

      // Clear results after successful replacement
      if (successful.length > 0) {
        resultsSection.classList.remove('visible');
        searchResults = [];
      }
    }
  } catch (error) {
    showStatus(`Error during replacement: ${error.message}`, 'error');
  } finally {
    replaceLoading.classList.remove('visible');
    updateReplaceButton();
  }
});
