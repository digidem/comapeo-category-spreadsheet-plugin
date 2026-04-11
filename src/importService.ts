/**
 * CoMapeo Config Import Service v2.0.0
 * Handles importing .comapeocat files and populating spreadsheet
 */

// =============================================================================
// Constants
// =============================================================================

/** Minimum valid file size for .comapeocat files (bytes) */
const MIN_COMAPEOCAT_SIZE = 100;

/** Maximum file size to prevent memory issues (10MB) */
const MAX_COMAPEOCAT_SIZE = 10 * 1024 * 1024;

/** ZIP file signature bytes */
const ZIP_SIGNATURE = [0x50, 0x4B, 0x03, 0x04];

// =============================================================================
// Import Dialog HTML Builder
// =============================================================================

/**
 * Builds the CSS styles for the import dialog.
 * Extracted for maintainability and readability.
 */
function buildImportDialogStyles(): string {
  return `
    * { box-sizing: border-box; }
    body {
      font-family: 'Roboto', Arial, sans-serif;
      margin: 0;
      padding: 20px;
      background: linear-gradient(135deg, #1a1a1a, #2c2c2c);
      color: #e0e0e0;
    }
    h2 {
      color: #6d44d9;
      font-size: 18px;
      margin: 0 0 16px 0;
      text-align: center;
    }
    .section {
      background: rgba(255, 255, 255, 0.05);
      border-radius: 10px;
      padding: 16px;
      margin-bottom: 16px;
    }
    .section-title {
      font-weight: 500;
      color: #9c88ff;
      margin-bottom: 10px;
      font-size: 13px;
    }
    .upload-area {
      border: 2px dashed #6d44d9;
      border-radius: 8px;
      padding: 24px 16px;
      text-align: center;
      cursor: pointer;
      transition: all 0.3s ease;
      background: rgba(109, 68, 217, 0.1);
    }
    .upload-area:hover, .upload-area.dragover {
      background: rgba(109, 68, 217, 0.2);
      border-color: #8a67e8;
    }
    .upload-area.has-file {
      border-color: #4CAF50;
      background: rgba(76, 175, 80, 0.1);
    }
    .upload-area.importing {
      pointer-events: none;
      opacity: 0.7;
    }
    .upload-icon { font-size: 32px; margin-bottom: 8px; }
    .upload-text { font-size: 13px; color: #aaa; }
    .upload-text strong { color: #6d44d9; }
    .file-input { display: none; }
    .file-info {
      margin-top: 12px;
      padding: 10px;
      background: rgba(0,0,0,0.3);
      border-radius: 6px;
      display: none;
    }
    .file-info.show { display: block; }
    .file-name { font-weight: 500; color: #fff; word-break: break-all; font-size: 13px; }
    .file-details { font-size: 11px; color: #888; margin-top: 4px; }
    .format-badge {
      display: inline-block;
      padding: 2px 6px;
      border-radius: 3px;
      font-size: 10px;
      font-weight: 500;
      margin-left: 6px;
    }
    .format-comapeocat { background: #6d44d9; color: white; }
    .format-mapeosettings { background: #ff9800; color: white; }
    .format-zip { background: #2196f3; color: white; }
    .format-unknown { background: #f44336; color: white; }
    .divider {
      text-align: center;
      color: #666;
      margin: 16px 0;
      position: relative;
      font-size: 12px;
    }
    .divider::before, .divider::after {
      content: '';
      position: absolute;
      top: 50%;
      width: 40%;
      height: 1px;
      background: #444;
    }
    .divider::before { left: 0; }
    .divider::after { right: 0; }
    .drive-input {
      width: 100%;
      padding: 10px;
      border: 1px solid #444;
      border-radius: 6px;
      background: rgba(0,0,0,0.3);
      color: #e0e0e0;
      font-size: 13px;
    }
    .drive-input:focus {
      outline: none;
      border-color: #6d44d9;
    }
    .drive-input:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }
    .info-text { font-size: 10px; color: #888; margin-top: 6px; }
    .btn {
      display: block;
      width: 100%;
      padding: 12px 20px;
      background: linear-gradient(45deg, #330B9E, #6d44d9);
      color: white;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 500;
      cursor: pointer;
      transition: all 0.3s ease;
      margin-top: 16px;
    }
    .btn:hover:not(:disabled) {
      background: linear-gradient(45deg, #4A0ED6, #8a67e8);
    }
    .btn:disabled {
      background: #444;
      cursor: not-allowed;
    }
    .error-msg {
      color: #ff6b6b;
      font-size: 12px;
      margin-top: 8px;
      padding: 8px;
      background: rgba(255, 107, 107, 0.1);
      border-radius: 6px;
      display: none;
    }
    .error-msg.show { display: block; }
    .progress-container {
      margin-top: 16px;
      display: none;
    }
    .progress-container.show { display: block; }
    .progress-header {
      display: flex;
      justify-content: space-between;
      margin-bottom: 6px;
      font-size: 12px;
      color: #aaa;
    }
    .progress-bar-bg {
      width: 100%;
      height: 8px;
      background: rgba(255,255,255,0.1);
      border-radius: 4px;
      overflow: hidden;
    }
    .progress-bar {
      height: 100%;
      background: linear-gradient(90deg, #330B9E, #6d44d9, #8a67e8);
      border-radius: 4px;
      transition: width 0.3s ease;
      width: 0%;
    }
    .progress-detail {
      margin-top: 6px;
      font-size: 11px;
      color: #888;
      text-align: center;
      min-height: 16px;
    }
    .spinner {
      display: inline-block;
      width: 14px;
      height: 14px;
      border: 2px solid #ffffff40;
      border-top-color: #fff;
      border-radius: 50%;
      animation: spin 0.8s linear infinite;
      margin-right: 8px;
      vertical-align: middle;
    }
    @keyframes spin { to { transform: rotate(360deg); } }
    .supported-formats {
      font-size: 10px;
      color: #666;
      text-align: center;
      margin-top: 10px;
    }
    .supported-formats code {
      background: rgba(109, 68, 217, 0.2);
      padding: 1px 4px;
      border-radius: 3px;
      color: #9c88ff;
    }
  `;
}

/**
 * Builds the HTML body for the import dialog.
 */
function buildImportDialogBody(): string {
  return `
    <h2>Import Category File</h2>

    <div class="section">
      <div class="section-title">📁 Upload from Computer</div>
      <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
        <div class="upload-icon">📤</div>
        <div class="upload-text">
          <strong>Click to select</strong> or drag and drop
        </div>
      </div>
      <input type="file" id="fileInput" class="file-input"
             accept=".comapeocat,.mapeosettings,.zip,.tar"
             onchange="handleFileSelect(event)">
      <div id="fileInfo" class="file-info">
        <span class="file-name" id="fileName"></span>
        <span class="format-badge" id="formatBadge"></span>
        <div class="file-details" id="fileDetails"></div>
      </div>
    </div>

    <div class="divider">or</div>

    <div class="section">
      <div class="section-title">☁️ Import from Google Drive</div>
      <input type="text" id="driveInput" class="drive-input"
             placeholder="Paste file ID or Drive URL"
             onkeypress="if(event.key==='Enter')handleImport()">
      <div class="info-text">
        File ID from URL: drive.google.com/file/d/<strong>FILE_ID</strong>/view
      </div>
    </div>

    <div id="progressContainer" class="progress-container">
      <div class="progress-header">
        <span id="progressStage">Processing...</span>
        <span id="progressPercent">0%</span>
      </div>
      <div class="progress-bar-bg">
        <div id="progressBar" class="progress-bar"></div>
      </div>
      <div id="progressDetail" class="progress-detail"></div>
    </div>

    <div id="errorMsg" class="error-msg"></div>

    <button id="importBtn" class="btn" onclick="handleImport()" disabled>
      Select a file to import
    </button>

    <div class="supported-formats">
      Supported: <code>.comapeocat</code> <code>.mapeosettings</code> <code>.zip</code>
    </div>
  `;
}

/**
 * Builds the JavaScript for the import dialog.
 * Note: Real-time progress updates from server to client are not possible in Apps Script.
 * We show an indeterminate progress indicator instead.
 */
function buildImportDialogScript(): string {
  return `
    let selectedFile = null;
    let importSource = null;
    let isImporting = false;
    const uploadArea = document.getElementById('uploadArea');

    // Drag and drop handlers
    uploadArea.addEventListener('dragover', function(e) {
      e.preventDefault();
      if (!isImporting) uploadArea.classList.add('dragover');
    });
    uploadArea.addEventListener('dragleave', function() {
      uploadArea.classList.remove('dragover');
    });
    uploadArea.addEventListener('drop', function(e) {
      e.preventDefault();
      uploadArea.classList.remove('dragover');
      if (!isImporting && e.dataTransfer.files.length > 0) {
        handleFile(e.dataTransfer.files[0]);
      }
    });

    // Drive input handler
    document.getElementById('driveInput').addEventListener('input', function() {
      if (isImporting) return;
      if (this.value.trim()) {
        importSource = 'drive';
        selectedFile = null;
        updateUI();
      } else if (!selectedFile) {
        importSource = null;
        updateUI();
      }
    });

    function handleFileSelect(event) {
      if (!isImporting && event.target.files[0]) {
        handleFile(event.target.files[0]);
      }
    }

    function handleFile(file) {
      selectedFile = file;
      importSource = 'local';
      document.getElementById('driveInput').value = '';
      document.getElementById('fileName').textContent = file.name;
      document.getElementById('fileDetails').textContent = 'Size: ' + formatFileSize(file.size);

      const format = detectFormat(file.name);
      const badge = document.getElementById('formatBadge');
      badge.textContent = formatToLabel(format);
      badge.className = 'format-badge format-' + format;

      document.getElementById('fileInfo').classList.add('show');
      uploadArea.classList.add('has-file');
      updateUI();
    }

    function detectFormat(filename) {
      const lower = filename.toLowerCase();
      if (lower.endsWith('.comapeocat')) return 'comapeocat';
      if (lower.endsWith('.mapeosettings') || lower.endsWith('.tar')) return 'mapeosettings';
      if (lower.endsWith('.zip')) return 'zip';
      return 'unknown';
    }

    function formatToLabel(format) {
      switch(format) {
        case 'comapeocat': return 'CoMapeo';
        case 'mapeosettings': return 'Legacy Mapeo';
        case 'zip': return 'ZIP';
        default: return 'Unknown';
      }
    }

    function formatFileSize(bytes) {
      if (bytes < 1024) return bytes + ' B';
      if (bytes < 1024*1024) return (bytes/1024).toFixed(1) + ' KB';
      return (bytes/1024/1024).toFixed(1) + ' MB';
    }

    function updateUI() {
      const btn = document.getElementById('importBtn');
      const driveInput = document.getElementById('driveInput');

      if (isImporting) {
        btn.disabled = true;
        btn.innerHTML = '<span class="spinner"></span>Importing...';
        driveInput.disabled = true;
        uploadArea.classList.add('importing');
        return;
      }

      driveInput.disabled = false;
      uploadArea.classList.remove('importing');

      const driveValue = driveInput.value.trim();
      if (importSource === 'local' && selectedFile) {
        btn.disabled = false;
        btn.innerHTML = 'Import File';
      } else if (importSource === 'drive' && driveValue) {
        btn.disabled = false;
        btn.innerHTML = 'Import from Drive';
      } else {
        btn.disabled = true;
        btn.innerHTML = 'Select a file to import';
      }
    }

    function showProgress(show) {
      const container = document.getElementById('progressContainer');
      if (show) {
        container.classList.add('show');
      } else {
        container.classList.remove('show');
      }
    }

    function updateProgress(data) {
      document.getElementById('progressBar').style.width = data.percent + '%';
      document.getElementById('progressStage').textContent = data.stage || 'Processing...';
      document.getElementById('progressPercent').textContent = data.percent + '%';

      let detail = data.detail || '';
      if (data.counts) {
        const counts = Object.entries(data.counts)
          .map(function(entry) { return entry[0] + ': ' + entry[1]; })
          .join(', ');
        detail += (detail ? ' ' : '') + '(' + counts + ')';
      }
      document.getElementById('progressDetail').textContent = detail;
    }

    function handleImport() {
      if (isImporting) return;

      isImporting = true;
      const errorMsg = document.getElementById('errorMsg');
      errorMsg.classList.remove('show');
      updateUI();

      if (importSource === 'local' && selectedFile) {
        showProgress(true);
        updateProgress({ percent: 0, stage: 'Reading file...', detail: '' });

        const reader = new FileReader();
        reader.onload = function(e) {
          // Show indeterminate progress while server processes
          updateProgress({ percent: 50, stage: 'Processing...', detail: 'Importing configuration' });
          const base64data = e.target.result.split(',')[1];

          // Use processImportedCategoryFile - progress callbacks cannot work
          // across Apps Script client/server boundary
          google.script.run
            .withSuccessHandler(handleSuccess)
            .withFailureHandler(handleError)
            .processImportedCategoryFile(selectedFile.name, base64data);
        };
        reader.onerror = function() {
          handleError({message: 'Failed to read file'});
        };
        reader.readAsDataURL(selectedFile);
      } else if (importSource === 'drive') {
        showProgress(true);
        updateProgress({ percent: 0, stage: 'Fetching from Drive...', detail: '' });

        const driveInput = document.getElementById('driveInput').value.trim();
        google.script.run
          .withSuccessHandler(handleSuccess)
          .withFailureHandler(handleError)
          .processImportFile(driveInput);
      }
    }

    function handleSuccess(result) {
      isImporting = false;
      showProgress(false);

      if (result && result.success === false) {
        handleError({message: result.message || 'Import failed'});
        return;
      }

      google.script.host.close();
    }

    function handleError(error) {
      isImporting = false;
      showProgress(false);
      updateUI();

      const errorMsg = document.getElementById('errorMsg');
      errorMsg.textContent = error.message || String(error);
      errorMsg.classList.add('show');
    }
  `;
}

/**
 * Builds the complete HTML for the import dialog.
 */
function buildImportDialogHtml(): string {
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>${buildImportDialogStyles()}</style>
</head>
<body>
  ${buildImportDialogBody()}
  <script>${buildImportDialogScript()}</script>
</body>
</html>`;
}

// =============================================================================
// Main Import Functions
// =============================================================================

/**
 * Prompts user to select a category file and imports it.
 * Supports local file upload and Google Drive import.
 * Accepts .comapeocat, .mapeosettings, and .zip files.
 */
function importCoMapeoCatFile(): void {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(buildImportDialogHtml())
    .setWidth(480)
    .setHeight(560);

  ui.showModalDialog(html, 'Import Category File');
}

/**
 * Processes the import file from Google Drive.
 * Supports .comapeocat, .mapeosettings, and .zip files.
 * @param fileIdOrUrl - Google Drive file ID or URL
 */
function processImportFile(fileIdOrUrl: string): { success: boolean; message: string } {
  // Validate input
  if (!fileIdOrUrl || typeof fileIdOrUrl !== 'string') {
    throw new Error('Please provide a valid file ID or URL.');
  }

  const trimmedInput = fileIdOrUrl.trim();
  if (trimmedInput.length === 0) {
    throw new Error('Please provide a valid file ID or URL.');
  }

  const fileId = extractFileId(trimmedInput);

  if (!fileId) {
    throw new Error('Invalid file ID or URL format. Please provide a valid Google Drive file ID or URL.');
  }

  // Validate file ID format (alphanumeric with hyphens/underscores)
  if (!/^[\w-]+$/.test(fileId)) {
    throw new Error('Invalid file ID format.');
  }

  let file: GoogleAppsScript.Drive.File;
  try {
    file = DriveApp.getFileById(fileId);
  } catch (e) {
    throw new Error('Could not access file. Please check the file ID and ensure you have permission to access it.');
  }

  // Validate file size
  const fileSize = file.getSize();
  if (fileSize < MIN_COMAPEOCAT_SIZE) {
    throw new Error('File is too small to be a valid category file.');
  }
  if (fileSize > MAX_COMAPEOCAT_SIZE) {
    throw new Error('File is too large (max 10MB).');
  }

  let blob: GoogleAppsScript.Base.Blob;
  try {
    blob = file.getBlob();
  } catch (e) {
    throw new Error('Could not read file contents. The file may be corrupted or inaccessible.');
  }

  // Convert blob to base64 and use the unified import pipeline
  const base64Data = Utilities.base64Encode(blob.getBytes());
  const fileName = file.getName();

  // Delegate to processImportedCategoryFile (defined in src/importCategory.ts)
  // which handles all formats: .comapeocat, .mapeosettings, .zip
  const result = processImportedCategoryFile(fileName, base64Data);

  if (!result.success) {
    throw new Error(result.message);
  }

  // Note: No UI alert here - the HTML dialog handles success/failure display
  // to maintain consistent behavior between local and Drive imports
  return result;
}

interface BuildValidationOptions {
  strict?: boolean;
}

/**
 * Validates that a BuildRequest has required fields
 */
function validateBuildRequest(content: BuildRequest, options?: BuildValidationOptions): void {
  if (!content) {
    throw new Error('Invalid configuration: empty content.');
  }

  if (!content.metadata) {
    throw new Error('Invalid configuration: missing metadata.');
  }

  if (!content.metadata.name || typeof content.metadata.name !== 'string') {
    throw new Error('Invalid configuration: missing or invalid metadata.name.');
  }

  if (!content.metadata.version || typeof content.metadata.version !== 'string') {
    throw new Error('Invalid configuration: missing or invalid metadata.version.');
  }

  if (!Array.isArray(content.categories)) {
    throw new Error('Invalid configuration: categories must be an array.');
  }

  if (!Array.isArray(content.fields)) {
    throw new Error('Invalid configuration: fields must be an array.');
  }

  if (options?.strict) {
    runStrictBuildValidations(content);
  }
}

function runStrictBuildValidations(buildRequest: BuildRequest): void {
  const log = getScopedLogger("ImportService");
  const errors: string[] = [];
  const warnings: string[] = [];

  const metadataName = String(buildRequest.metadata?.name || '').trim();
  if (!metadataName) {
    errors.push('metadata.name must be a non-empty string.');
  } else if (containsUnsafeNameCharacters(metadataName)) {
    errors.push('metadata.name cannot contain slashes, backslashes, or ellipses.');
  }

  const metadataVersion = buildRequest.metadata?.version;
  if (metadataVersion) {
    const versionStr = String(metadataVersion).trim();
    if (!versionStr) {
      errors.push('metadata.version must not be blank when provided.');
    } else if (containsUnsafeNameCharacters(versionStr)) {
      errors.push('metadata.version cannot contain slashes, backslashes, or ellipses.');
    }
  }

  const categories = Array.isArray(buildRequest.categories) ? buildRequest.categories : [];
  if (categories.length === 0) {
    errors.push('At least one category is required.');
  }

  const fields = Array.isArray(buildRequest.fields) ? buildRequest.fields : [];
  if (fields.length === 0) {
    errors.push('At least one field is required.');
  }

  const allowedApplies = new Set(['observation', 'track']);
  const normalizedFieldIds = new Set<string>();
  const normalizedFieldIdsLower = new Set<string>();
  const allowedFieldTypes = new Set([
    'text',
    'textarea',
    'number',
    'integer',
    'select',
    'selectone',
    'multiselect',
    'selectmultiple',
    'boolean',
    'date',
    'datetime',
    'photo',
    'location'
  ]);
  const selectLikeTypes = new Set(['select', 'selectone', 'multiselect', 'selectmultiple']);

  fields.forEach((field, index) => {
    const fieldLabel = `Field ${field?.id || index + 1}`;
    if (!field || typeof field !== 'object') {
      errors.push(`${fieldLabel} is invalid.`);
      return;
    }

    const fieldId = String(field.id || '').trim();
    if (!fieldId) {
      errors.push(`${fieldLabel} must have an id.`);
    } else {
      const lower = fieldId.toLowerCase();
      if (normalizedFieldIdsLower.has(lower)) {
        errors.push(`Duplicate field id detected: "${fieldId}".`);
      }
      normalizedFieldIds.add(fieldId);
      normalizedFieldIdsLower.add(lower);
    }

    const normalizedType = normalizeFieldTypeForValidation(field.type);
    if (!normalizedType) {
      errors.push(`${fieldLabel} must have a valid type.`);
      return;
    }

    if (!allowedFieldTypes.has(normalizedType)) {
      errors.push(`${fieldLabel} has unsupported type "${field.type}".`);
    }

    if (selectLikeTypes.has(normalizedType)) {
      if (!Array.isArray(field.options) || field.options.length === 0) {
        errors.push(`${fieldLabel} (${field.type}) must include at least one option.`);
      }
    }
  });

  let hasTrackCategory = false;
  const ensureFieldReferencesAreValid = (ids: string[] | undefined | null, label: string) => {
    if (!ids || !Array.isArray(ids)) return;
    ids.forEach(refId => {
      if (!refId) return;
      const normalizedRef = String(refId).trim();
      if (!normalizedFieldIds.has(normalizedRef)) {
        errors.push(`${label} references unknown field id "${refId}".`);
      }
    });
  };

  categories.forEach((category, index) => {
    const categoryLabel = `Category ${category?.id || index + 1}`;
    if (!category || typeof category !== 'object') {
      errors.push(`${categoryLabel} is invalid.`);
      return;
    }

    const categoryId = String(category.id || '').trim();
    if (!categoryId) {
      errors.push(`${categoryLabel} must have an id.`);
    }

    const categoryName = String(category.name || '').trim();
    if (!categoryName) {
      errors.push(`${categoryLabel} must have a name.`);
    }

    const appliesTo = category.appliesTo;
    if (!Array.isArray(appliesTo) || appliesTo.length === 0) {
      errors.push(`${categoryLabel} must include appliesTo values (observation and/or track).`);
    } else {
      const normalizedApplies = appliesTo.map(value => String(value || '').trim().toLowerCase()).filter(Boolean);
      const invalidApply = normalizedApplies.find(value => !allowedApplies.has(value));
      if (invalidApply) {
        errors.push(`${categoryLabel} has invalid appliesTo value "${invalidApply}".`);
      }
      if (normalizedApplies.includes('track')) {
        hasTrackCategory = true;
      }
    }

    ensureFieldReferencesAreValid(category.defaultFieldIds, `${categoryLabel} defaultFieldIds`);
    ensureFieldReferencesAreValid(category.fields, `${categoryLabel} fields`);
  });

  if (!hasTrackCategory) {
    warnings.push('No categories are marked with "track". The track viewer will be empty unless at least one category includes "track" in appliesTo.');
  }

  const icons = Array.isArray(buildRequest.icons) ? buildRequest.icons : [];
  const iconIds = new Set<string>();
  const MAX_INLINE_SVG_BYTES = 2 * 1024 * 1024; // 2 MB
  const LARGE_SVG_WARNING_BYTES = 300 * 1024; // 300 KB

  icons.forEach((icon, index) => {
    const iconLabel = `Icon ${icon?.id || index + 1}`;
    if (!icon || typeof icon !== 'object') {
      errors.push(`${iconLabel} is invalid.`);
      return;
    }

    const iconId = String(icon.id || '').trim();
    if (!iconId) {
      errors.push(`${iconLabel} must have an id.`);
    } else if (iconIds.has(iconId)) {
      errors.push(`Duplicate icon id detected: "${iconId}".`);
    } else {
      iconIds.add(iconId);
    }

    const hasSvgData = typeof icon.svgData === 'string' && icon.svgData.trim().length > 0;
    const hasSvgUrl = typeof icon.svgUrl === 'string' && icon.svgUrl.trim().length > 0;
    if (!hasSvgData && !hasSvgUrl) {
      errors.push(`${iconLabel} must include svgData or svgUrl.`);
    }

    if (hasSvgData) {
      const svgBytes = getByteLength(icon.svgData || '');
      if (svgBytes > MAX_INLINE_SVG_BYTES) {
        errors.push(`${iconLabel} inline SVG exceeds 2 MB (${Math.round(svgBytes / 1024)} KB).`);
      } else if (svgBytes > LARGE_SVG_WARNING_BYTES) {
        warnings.push(`${iconLabel} inline SVG is ${Math.round(svgBytes / 1024)} KB. Consider optimizing large assets.`);
      }
    }
  });

  const translations = buildRequest.translations;
  const MAX_LOCALE_BYTES = 1 * 1024 * 1024; // 1 MB
  if (translations && typeof translations === 'object') {
    Object.keys(translations).forEach(locale => {
      if (!isValidBcp47Locale(locale)) {
        errors.push(`Translation locale "${locale}" must be a valid BCP-47 tag (use hyphens, not underscores).`);
      }

      const localeData = translations[locale];
      const serialized = JSON.stringify(localeData || {});
      const localeBytes = getByteLength(serialized);
      if (localeBytes > MAX_LOCALE_BYTES) {
        errors.push(`Translations for locale "${locale}" exceed 1 MB (${Math.round(localeBytes / 1024)} KB).`);
      }
    });
  }

  const optionCount = fields.reduce((total, field) => total + (Array.isArray(field.options) ? field.options.length : 0), 0);
  const translationEntryCount = countTranslationEntries(translations);
  const totalEntities = categories.length + fields.length + icons.length + optionCount + translationEntryCount;
  const MAX_TOTAL_ENTITIES = 10000;
  if (totalEntities > MAX_TOTAL_ENTITIES) {
    errors.push(`Combined counts (categories + fields + icons + options + translations) exceed ${MAX_TOTAL_ENTITIES}. Current total: ${totalEntities}.`);
  }

  if (warnings.length > 0) {
    warnings.forEach(warning => log.warn(`[Build Validation Warning] ${warning}`));
  }

  if (errors.length > 0) {
    throw new Error(`Build payload validation failed:\n - ${errors.join('\n - ')}`);
  }
}

function containsUnsafeNameCharacters(value: string): boolean {
  return /[\\/]/.test(value) || value.includes('...');
}

function normalizeFieldTypeForValidation(type: any): string {
  if (!type && type !== 0) return '';
  const normalized = String(type).trim().toLowerCase();
  const collapsed = normalized.replace(/[^a-z0-9]/g, '');
  switch (collapsed) {
    case 'selectone':
    case 'single':
      return 'select';
    case 'selectmultiple':
    case 'multiselect':
    case 'multi':
      return 'multiselect';
    default:
      return normalized;
  }
}

function isValidBcp47Locale(locale: string): boolean {
  if (!locale || typeof locale !== 'string') return false;
  if (locale.includes('_')) return false;
  const trimmed = locale.trim();
  if (!trimmed) return false;
  const bcp47Regex = /^[a-z]{2,3}(-[a-z0-9]{2,8})*$/i;
  return bcp47Regex.test(trimmed);
}

function countTranslationEntries(translations: any): number {
  if (!translations || typeof translations !== 'object') {
    return 0;
  }

  let total = 0;
  Object.keys(translations).forEach(locale => {
    const entry = translations[locale];
    if (!entry || typeof entry !== 'object') {
      return;
    }

    const metadata = entry.metadata;
    if (metadata && typeof metadata === 'object') {
      total += Object.values(metadata).filter(value => typeof value === 'string' && value).length;
    }

    const categories = entry.categories || entry.category;
    if (categories && typeof categories === 'object') {
      Object.keys(categories).forEach(catId => {
        const catEntry = categories[catId];
        if (!catEntry || typeof catEntry !== 'object') return;
        total += Object.values(catEntry).filter(value => typeof value === 'string' && value).length;
      });
    }

    const fields = entry.fields || entry.field;
    if (fields && typeof fields === 'object') {
      Object.keys(fields).forEach(fieldId => {
        const fieldEntry = fields[fieldId];
        if (!fieldEntry || typeof fieldEntry !== 'object') return;
        total += Object.values(fieldEntry).filter(value => typeof value === 'string' && value).length;
      });
    }
  });

  return total;
}

function getByteLength(value: string): number {
  const log = getScopedLogger("ImportService");
  const str = value || '';
  try {
    if (typeof Utilities !== 'undefined' && Utilities && typeof Utilities.newBlob === 'function') {
      return Utilities.newBlob(str).getBytes().length;
    }
  } catch (error) {
    log.warn('Unable to calculate byte length via Utilities:', error);
  }

  try {
    const maybeBuffer = (globalThis as any)?.Buffer;
    if (typeof maybeBuffer?.byteLength === 'function') {
      return maybeBuffer.byteLength(str, 'utf8');
    }
  } catch (error) {
    log.warn('Unable to calculate byte length via Buffer:', error);
  }

  let bytes = 0;
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    if (code <= 0x7f) bytes += 1;
    else if (code <= 0x7ff) bytes += 2;
    else if (code >= 0xd800 && code <= 0xdfff) {
      bytes += 4;
      i++;
    } else if (code < 0xffff) bytes += 3;
    else bytes += 4;
  }
  return bytes;
}

/**
 * Extracts file ID from a Drive URL or returns the ID if already provided
 */
function extractFileId(input: string): string | null {
  if (!input || typeof input !== 'string') return null;

  const trimmed = input.trim();
  if (trimmed.length === 0) return null;

  // Check if it's already a file ID (alphanumeric with hyphens/underscores, reasonable length)
  if (/^[\w-]{10,100}$/.test(trimmed)) {
    return trimmed;
  }

  // Extract from various Google Drive URL formats
  const patterns = [
    /\/file\/d\/([^\/\?]+)/,   // /file/d/FILE_ID
    /[?&]id=([^&]+)/,          // ?id=FILE_ID or &id=FILE_ID
    /\/d\/([^\/\?]+)/,         // /d/FILE_ID
    /\/open\?id=([^&]+)/       // /open?id=FILE_ID
  ];

  for (const pattern of patterns) {
    const match = trimmed.match(pattern);
    if (match && match[1]) {
      const extractedId = match[1].trim();
      // Validate extracted ID format
      if (/^[\w-]{10,100}$/.test(extractedId)) {
        return extractedId;
      }
    }
  }

  return null;
}

// =============================================================================
// File Parsing
// =============================================================================

/**
 * Extracts configuration JSON from a .comapeocat file
 * .comapeocat files are ZIP archives containing configuration data
 */
function extractConfigFromComapeocat(blob: GoogleAppsScript.Base.Blob): BuildRequest {
  if (!blob) {
    throw new Error('Invalid file: no data.');
  }

  let bytes: number[];
  try {
    bytes = blob.getBytes();
  } catch (e) {
    throw new Error('Could not read file bytes.');
  }

  if (!bytes || bytes.length < MIN_COMAPEOCAT_SIZE) {
    throw new Error('File is too small or empty.');
  }

  // Check for ZIP signature (PK\x03\x04)
  const isZip = bytes.length >= 4 &&
    bytes[0] === ZIP_SIGNATURE[0] &&
    bytes[1] === ZIP_SIGNATURE[1] &&
    bytes[2] === ZIP_SIGNATURE[2] &&
    bytes[3] === ZIP_SIGNATURE[3];

  if (!isZip) {
    // Try parsing as raw JSON (for backwards compatibility or direct JSON export)
    return parseAsJson(blob);
  }

  // Unzip and find configuration
  return parseAsZip(blob);
}

/**
 * Attempts to parse blob as raw JSON
 */
function parseAsJson(blob: GoogleAppsScript.Base.Blob): BuildRequest {
  let jsonStr: string;
  try {
    jsonStr = blob.getDataAsString('UTF-8');
  } catch (e) {
    throw new Error('Could not read file as text.');
  }

  if (!jsonStr || jsonStr.trim().length === 0) {
    throw new Error('File is empty.');
  }

  let parsed: any;
  try {
    parsed = JSON.parse(jsonStr);
  } catch (e) {
    throw new Error('Invalid file format: not a valid JSON or ZIP file.');
  }

  // Validate it has expected structure
  if (!parsed || typeof parsed !== 'object') {
    throw new Error('Invalid JSON: expected an object.');
  }

  if (!parsed.metadata || !Array.isArray(parsed.categories)) {
    throw new Error('Invalid configuration structure: missing metadata or categories.');
  }

  return parsed as BuildRequest;
}

/**
 * Parses a ZIP blob to extract configuration
 */
function parseAsZip(blob: GoogleAppsScript.Base.Blob): BuildRequest {
  let unzipped: GoogleAppsScript.Base.Blob[];
  try {
    unzipped = Utilities.unzip(blob);
  } catch (e) {
    throw new Error('Could not unzip file. The file may be corrupted.');
  }

  if (!unzipped || unzipped.length === 0) {
    throw new Error('ZIP file is empty.');
  }

  // Look for config.json first
  for (const file of unzipped) {
    const name = file.getName();
    if (name === 'config.json' || name.endsWith('/config.json')) {
      return parseJsonFile(file, name);
    }
  }

  // Try to find any JSON file with the expected structure
  for (const file of unzipped) {
    const name = file.getName();
    if (name.endsWith('.json')) {
      try {
        const content = parseJsonFile(file, name);
        if (content.metadata && Array.isArray(content.categories)) {
          return content;
        }
      } catch {
        // Continue to next file
        continue;
      }
    }
  }

  throw new Error('Could not find configuration data in .comapeocat file. Expected config.json or similar.');
}

/**
 * Parses a single JSON file from the ZIP
 */
function parseJsonFile(file: GoogleAppsScript.Base.Blob, fileName: string): BuildRequest {
  let content: string;
  try {
    content = file.getDataAsString('UTF-8');
  } catch (e) {
    throw new Error(`Could not read ${fileName}.`);
  }

  let parsed: any;
  try {
    parsed = JSON.parse(content);
  } catch (e) {
    throw new Error(`Invalid JSON in ${fileName}.`);
  }

  return parsed as BuildRequest;
}

// =============================================================================
// Spreadsheet Population
// =============================================================================

/**
 * Builds a lookup map from icon ID to icon data (svgData or svgUrl)
 * This allows us to persist actual icon content when importing categories
 */
function buildIconMap(icons?: Icon[]): Map<string, string> {
  const map = new Map<string, string>();

  if (!icons || !Array.isArray(icons)) return map;

  for (const icon of icons) {
    if (!icon || !icon.id) continue;

    // Prefer svgData (inline SVG), fall back to svgUrl
    if (icon.svgData) {
      // If it's raw SVG, convert to data URI for storage
      if (icon.svgData.startsWith('<svg')) {
        map.set(icon.id, `data:image/svg+xml,${encodeURIComponent(icon.svgData)}`);
      } else {
        map.set(icon.id, icon.svgData);
      }
    } else if (icon.svgUrl) {
      map.set(icon.id, icon.svgUrl);
    }
  }

  return map;
}

/**
 * Populates the spreadsheet with imported configuration data
 */
function populateSpreadsheetFromConfig(config: BuildRequest): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error('No active spreadsheet found.');
  }

  // Set category selection from imported order
  if (config.categories && config.categories.length > 0) {
    setCategorySelection(config.categories.map(c => c.id));
  }

  // Build icon lookup map for populating categories with actual icon data
  const iconMap = buildIconMap(config.icons);

  // Populate sheets in order
  // NOTE: We do NOT auto-detect universal fields during import to preserve round-trip integrity.
  // Fields are placed exactly as they appear in the imported config's defaultFieldIds.
  populateCategoriesSheet(spreadsheet, config.categories || [], iconMap, config.fields);
  populateDetailsSheet(spreadsheet, config.fields || []);
  populateMetadataSheet(spreadsheet, config.metadata, config);
  populateIconsSheet(spreadsheet, config.icons);

  // Populate translations if present, otherwise clear existing translation sheets
  if (config.translations && Object.keys(config.translations).length > 0) {
    populateTranslationSheets(spreadsheet, config);
  } else {
    clearTranslationSheets(spreadsheet);
  }
}


/**
 * Populates the Categories sheet
 * @param iconMap - Map of icon ID to actual icon data (svgData URI or svgUrl)
 */
function populateCategoriesSheet(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  categories: Category[],
  iconMap: Map<string, string>,
  fields?: Field[]
): void {
  let sheet = spreadsheet.getSheetByName('Categories');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('Categories');
    sheet.getRange(1, 1, 1, 6).setValues([['Name', 'Icon', 'Fields', 'ID', 'Color', 'Icon ID']]).setFontWeight('bold');
  } else {
    // Clear existing data (keep header)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 6).clear();
    }
  }

  if (!categories || categories.length === 0) return;

  // Build map from field ID to field name for converting defaultFieldIds
  const fieldIdToName = new Map<string, string>();
  if (fields && fields.length > 0) {
    for (const field of fields) {
      if (field.id && field.name) {
        fieldIdToName.set(field.id, field.name);
      }
    }
  }

  const rows: any[][] = [];
  const colors: string[] = [];

  for (const cat of categories) {
    if (!cat || !cat.name) continue;

    // Look up actual icon data from the icon map using iconId
    // This ensures icons round-trip correctly (not just the ID)
    let iconValue = '';
    if (cat.iconId) {
      iconValue = iconMap.get(cat.iconId) || '';
    }

    // Convert field IDs to field names for column C
    // Preserve ALL fields exactly as they appear in the imported config (no filtering)
    let fieldsValue = '';
    const fieldIds = Array.isArray(cat.defaultFieldIds) && cat.defaultFieldIds.length > 0
      ? cat.defaultFieldIds
      : (Array.isArray((cat as any).fields) ? (cat as any).fields : []);

    if (fieldIds && fieldIds.length > 0) {
      const fieldNames = fieldIds
        .map(id => fieldIdToName.get(id) || id)  // Fall back to ID if name not found
        .filter(Boolean);
      fieldsValue = fieldNames.join(', ');
    }

    // Only set color if explicitly provided in the config
    // Leave empty to preserve undefined state (allows downstream defaults)
    const colorValue = cat.color || '';

    rows.push([cat.name, iconValue, fieldsValue, cat.id || '', colorValue, cat.iconId || '']);
    // Only set background color if we have a valid color value
    colors.push(colorValue || '#FFFFFF');  // Use white as background default for empty cells
  }

  if (rows.length === 0) return;

  sheet.getRange(2, 1, rows.length, 6).setValues(rows);

  // Set background colors for column A
  const bgColors = colors.map(c => [c]);
  sheet.getRange(2, 1, colors.length, 1).setBackgrounds(bgColors);
}

/**
 * Populates the Details sheet
 */
function populateDetailsSheet(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  fields: Field[]
): void {
  let sheet = spreadsheet.getSheetByName('Details');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('Details');
    sheet.getRange(1, 1, 1, 6).setValues([['Name', 'Helper Text', 'Type', 'Options', 'ID', 'Universal']]).setFontWeight('bold');
  } else {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 6).clear();
    }
  }

  if (!fields || fields.length === 0) return;

  const rows: any[][] = [];

  for (const field of fields) {
    if (!field || !field.name) continue;

    const typeChar = mapFieldTypeToChar(field.type);
    // Preserve option values by writing "value:label" when they differ
    let optionsValue = '';
    if (Array.isArray(field.options) && field.options.length > 0) {
      optionsValue = field.options.map(o => {
        const value = o?.value || '';
        const label = o?.label || '';
        if (!label) return '';
        // Only include value prefix if it differs from slugified label
        return value === slugify(label) ? label : `${value}:${label}`;
      }).filter(Boolean).join(', ');
    }

    // Never mark fields as Universal during import to preserve round-trip integrity.
    // The Universal flag is only set by users when creating/editing the spreadsheet.
    rows.push([
      field.name,
      field.description || '',
      typeChar,
      optionsValue,
      field.id || '',  // Store original field ID for round-trip
      'FALSE'  // Always FALSE on import to preserve exact field distribution
    ]);
  }

  if (rows.length === 0) return;

  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
}

/**
 * Maps API field type to spreadsheet type character
 */
function mapFieldTypeToChar(type: FieldType): string {
  switch (type) {
    case 'text':
      return 't';
    case 'textarea':
      return 'T';
    case 'number':
      return 'n';
    case 'integer':
      return 'i';
    case 'multiselect':
      return 'm';
    case 'select':
      return 's';
    case 'boolean':
      return 'b';
    case 'date':
      return 'd';
    case 'datetime':
      return 'D';
    case 'photo':
      return 'p';
    case 'location':
      return 'l';
    default:
      return 's';  // Fallback to select for unknown types
  }
}

/**
 * Detects the primary language from the config structure.
 * The primary language is the language in which category/field names are originally authored.
 * All other languages in translations are considered secondary.
 *
 * Strategy:
 * 1. If no translations exist, default to English
 * 2. Compare base names against translation locales to find which locale matches the base
 * 3. If no match, default to English
 */
function detectPrimaryLanguage(config: BuildRequest): string {
  // No translations means English default
  if (!config.translations || Object.keys(config.translations).length === 0) {
    return 'English';
  }

  const locales = Object.keys(config.translations);

  // Try to detect which locale the base content is authored in
  // by checking if any category names match their "translations"
  if (config.categories && config.categories.length > 0) {
    for (const locale of locales) {
      const categoryTranslations = config.translations[locale]?.categories;
      if (!categoryTranslations) continue;

      // Check if base names match this locale's translations
      // (indicating base content is in this language)
      let matchCount = 0;
      let totalChecked = 0;

      for (const cat of config.categories) {
        if (!cat || !cat.id) continue;
        const translation = categoryTranslations[cat.id]?.name;
        if (translation) {
          totalChecked++;
          if (cat.name === translation) {
            matchCount++;
          }
        }
      }

      // If majority of translations match base names, this is likely the primary language
      if (totalChecked > 0 && matchCount / totalChecked > 0.5) {
        // Use the language name from map if available, otherwise capitalize the locale code
        return ALL_LANGUAGES[locale] || locale.charAt(0).toUpperCase() + locale.slice(1);
      }
    }
  }

  // Default to English if no clear match
  return 'English';
}

/**
 * Populates the Metadata sheet, including primary language detection
 */
function populateMetadataSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, metadata: Metadata, config?: BuildRequest): void {
  if (!metadata) return;

  let sheet = spreadsheet.getSheetByName('Metadata');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('Metadata');
    sheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
  } else {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 2).clear();
    }
  }

  // Detect primary language from config structure
  const primaryLanguage = config ? detectPrimaryLanguage(config) : 'English';

  const rows = [
    ['name', metadata.name || ''],
    ['version', metadata.version || ''],
    ['description', metadata.description || ''],
    ['primaryLanguage', primaryLanguage]
  ];

  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

/**
 * Populates the Icons sheet with all icons from the config
 * This ensures all icons are preserved during round-trip, not just category icons
 */
function populateIconsSheet(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, icons?: Icon[]): void {
  let sheet = spreadsheet.getSheetByName('Icons');

  if (!sheet) {
    sheet = spreadsheet.insertSheet('Icons');
    sheet.getRange(1, 1, 1, 2).setValues([['ID', 'SVG Data']]).setFontWeight('bold');
  } else {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 2).clear();
    }
  }

  if (!icons || icons.length === 0) return;

  const rows: any[][] = [];
  for (const icon of icons) {
    if (!icon || !icon.id) continue;

    // Store either svgData or svgUrl
    const svgValue = icon.svgData || icon.svgUrl || '';
    rows.push([icon.id, svgValue]);
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

/**
 * Deletes all translation sheets to remove stale data when importing a config without translations.
 * Deleting rather than clearing allows auto-translate to recreate them with proper headers/formulas.
 */
function clearTranslationSheets(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet): void {
  const translationSheetNames = [
    'Category Translations',
    'Detail Label Translations',
    'Detail Helper Text Translations',
    'Detail Option Translations'
  ];

  for (const sheetName of translationSheetNames) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      spreadsheet.deleteSheet(sheet);
    }
  }
}

/**
 * Populates translation sheets from imported config
 */
function populateTranslationSheets(spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet, config: BuildRequest): void {
  if (!config.translations) return;

  const locales = Object.keys(config.translations).filter(Boolean);
  if (locales.length === 0) return;

  // Helper to clear entire sheet content (headers + data) to remove stale columns
  const clearSheetContent = (sheet: GoogleAppsScript.Spreadsheet.Sheet): void => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 0 && lastCol > 0) {
      sheet.getRange(1, 1, lastRow, lastCol).clear();
    }
  };

  // Create Category Translations sheet
  let catSheet = spreadsheet.getSheetByName('Category Translations');
  if (!catSheet) {
    catSheet = spreadsheet.insertSheet('Category Translations');
  } else {
    clearSheetContent(catSheet);
  }

  const catHeaders = ['Name', ...locales.map(l => `${l}`)];
  catSheet.getRange(1, 1, 1, catHeaders.length).setValues([catHeaders]).setFontWeight('bold');

  // Set column A formulas to reference Categories sheet (individual cell references)
  const categoriesSheet = spreadsheet.getSheetByName('Categories');
  if (categoriesSheet && config.categories && config.categories.length > 0) {
    const lastRow = categoriesSheet.getLastRow();
    if (lastRow > 1) {
      // Generate individual formulas for each row (setFormula only populates first cell)
      const formulas: string[][] = [];
      for (let i = 2; i <= lastRow; i++) {
        formulas.push([`=Categories!A${i}`]);
      }
      catSheet.getRange(2, 1, lastRow - 1, 1).setFormulas(formulas);
    }

    // Write translation values (columns B+)
    const catTransValues: any[][] = [];
    for (const cat of config.categories) {
      if (!cat || !cat.id) continue;
      const row: any[] = [];
      for (const locale of locales) {
        const trans = config.translations[locale]?.categories?.[cat.id]?.name || '';
        row.push(trans);
      }
      catTransValues.push(row);
    }
    if (catTransValues.length > 0 && locales.length > 0) {
      catSheet.getRange(2, 2, catTransValues.length, locales.length).setValues(catTransValues);
    }
  }

  // Create Detail Label Translations sheet
  let labelSheet = spreadsheet.getSheetByName('Detail Label Translations');
  if (!labelSheet) {
    labelSheet = spreadsheet.insertSheet('Detail Label Translations');
  } else {
    clearSheetContent(labelSheet);
  }

  const labelHeaders = ['Name', ...locales.map(l => `${l}`)];
  labelSheet.getRange(1, 1, 1, labelHeaders.length).setValues([labelHeaders]).setFontWeight('bold');

  // Set column A formulas to reference Details sheet column A (field names, individual cell references)
  const detailsSheet = spreadsheet.getSheetByName('Details');
  if (detailsSheet && config.fields && config.fields.length > 0) {
    const lastRow = detailsSheet.getLastRow();
    if (lastRow > 1) {
      // Generate individual formulas for each row (setFormula only populates first cell)
      const formulas: string[][] = [];
      for (let i = 2; i <= lastRow; i++) {
        formulas.push([`=Details!A${i}`]);
      }
      labelSheet.getRange(2, 1, lastRow - 1, 1).setFormulas(formulas);
    }

    // Write translation values (columns B+)
    const labelTransValues: any[][] = [];
    for (const field of config.fields) {
      if (!field || !field.id) continue;
      const row: any[] = [];
      for (const locale of locales) {
        const trans = config.translations[locale]?.fields?.[field.id]?.name || '';
        row.push(trans);
      }
      labelTransValues.push(row);
    }
    if (labelTransValues.length > 0 && locales.length > 0) {
      labelSheet.getRange(2, 2, labelTransValues.length, locales.length).setValues(labelTransValues);
    }
  }

  // Create Detail Helper Text Translations sheet
  let helperSheet = spreadsheet.getSheetByName('Detail Helper Text Translations');
  if (!helperSheet) {
    helperSheet = spreadsheet.insertSheet('Detail Helper Text Translations');
  } else {
    clearSheetContent(helperSheet);
  }

  const helperHeaders = ['Helper Text', ...locales.map(l => `${l}`)];
  helperSheet.getRange(1, 1, 1, helperHeaders.length).setValues([helperHeaders]).setFontWeight('bold');

  // Set column A formulas to reference Details sheet column B (helper text, individual cell references)
  if (detailsSheet && config.fields && config.fields.length > 0) {
    const lastRow = detailsSheet.getLastRow();
    if (lastRow > 1) {
      // Generate individual formulas for each row (setFormula only populates first cell)
      const formulas: string[][] = [];
      for (let i = 2; i <= lastRow; i++) {
        formulas.push([`=Details!B${i}`]);
      }
      helperSheet.getRange(2, 1, lastRow - 1, 1).setFormulas(formulas);
    }

    // Write translation values (columns B+)
    const helperTransValues: any[][] = [];
    for (const field of config.fields) {
      if (!field || !field.id) continue;
      const row: any[] = [];
      for (const locale of locales) {
        const trans = config.translations[locale]?.fields?.[field.id]?.description || '';
        row.push(trans);
      }
      helperTransValues.push(row);
    }
    if (helperTransValues.length > 0 && locales.length > 0) {
      helperSheet.getRange(2, 2, helperTransValues.length, locales.length).setValues(helperTransValues);
    }
  }

  // Create Detail Option Translations sheet
  let optionSheet = spreadsheet.getSheetByName('Detail Option Translations');
  if (!optionSheet) {
    optionSheet = spreadsheet.insertSheet('Detail Option Translations');
  } else {
    clearSheetContent(optionSheet);
  }

  const optionHeaders = ['Options', ...locales.map(l => `${l}`)];
  optionSheet.getRange(1, 1, 1, optionHeaders.length).setValues([optionHeaders]).setFontWeight('bold');

  // Set column A formulas to reference Details sheet column D (options, individual cell references)
  if (detailsSheet && config.fields && config.fields.length > 0) {
    const lastRow = detailsSheet.getLastRow();
    if (lastRow > 1) {
      // Generate individual formulas for each row (setFormula only populates first cell)
      const formulas: string[][] = [];
      for (let i = 2; i <= lastRow; i++) {
        formulas.push([`=Details!D${i}`]);
      }
      optionSheet.getRange(2, 1, lastRow - 1, 1).setFormulas(formulas);
    }

    // Write translation values (columns B+)
    const optionTransValues: any[][] = [];
    for (const field of config.fields) {
      if (!field || !field.id) continue;
      const row: any[] = [];
      for (const locale of locales) {
        const fieldTrans = config.translations[locale]?.fields?.[field.id];
        if (fieldTrans?.options && field.options) {
          // Join option translations in order of field.options
          const optTrans = field.options.map(opt => fieldTrans.options?.[opt.value] || '').join(', ');
          row.push(optTrans);
        } else {
          row.push('');
        }
      }
      optionTransValues.push(row);
    }
    if (optionTransValues.length > 0 && locales.length > 0) {
      optionSheet.getRange(2, 2, optionTransValues.length, locales.length).setValues(optionTransValues);
    }
  }
}
