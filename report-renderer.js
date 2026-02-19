// Get report data from the main process
window.electronAPI.getReportData().then(data => {
  renderReport(data);
});

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function renderReport(data) {
  const { searchText, replaceText, results } = data;

  // Render summary
  const totalReplacements = results.reduce((sum, r) => sum + (r.replacementCount || 0), 0);
  const summaryEl = document.getElementById('summary');
  summaryEl.innerHTML = `
    <p>
      <strong>${results.length}</strong> file(s) would be modified with
      <strong>${totalReplacements}</strong> total replacement(s).
      <br>
      Find: "<strong>${escapeHtml(searchText)}</strong>" →
      Replace: "<strong>${escapeHtml(replaceText)}</strong>"
    </p>
  `;

  // Render each file
  const contentEl = document.getElementById('reportContent');
  let html = '';

  for (const result of results) {
    const typeClass = result.fileType === 'Word' ? 'word' : 'powerpoint';

    html += `
      <div class="file-card">
        <div class="file-header">
          <span class="file-name">${escapeHtml(result.fileName)}</span>
          <span class="file-type ${typeClass}">${result.fileType}</span>
        </div>
        <div class="file-path">${escapeHtml(result.filePath)}</div>
        <div class="output-path">${escapeHtml(result.modifiedPath)}</div>
        <div class="snippets">
          <div class="snippet-count">${result.replacementCount} replacement(s) found</div>
    `;

    if (result.snippets && result.snippets.length > 0) {
      for (const snippet of result.snippets) {
        html += `
          <div class="snippet">
            <div class="snippet-label before">Before</div>
            <div class="snippet-text before">${formatSnippet(snippet.before, searchText, 'highlight')}</div>
            <div class="snippet-label after">After</div>
            <div class="snippet-text after">${formatSnippet(snippet.after, replaceText, 'replacement')}</div>
          </div>
        `;
      }
    } else {
      html += '<div class="no-snippets">No preview snippets available</div>';
    }

    html += `
        </div>
      </div>
    `;
  }

  contentEl.innerHTML = html;
}

function formatSnippet(text, term, className) {
  if (!term) return escapeHtml(text);

  // Escape HTML first
  const escaped = escapeHtml(text);
  const escapedTerm = escapeHtml(term);

  // Highlight the term (case-insensitive)
  const regex = new RegExp(`(${escapeRegex(escapedTerm)})`, 'gi');
  return escaped.replace(regex, `<span class="${className}">$1</span>`);
}

function escapeRegex(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
