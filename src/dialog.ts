/// <reference path="./lint.ts" />

/**
 * CoMapeo logo as embedded SVG data URI
 * Eliminates external network dependency and improves security
 */
const COMAPEO_LOGO_SVG = "data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20viewBox%3D%220%200%20100%20100%22%3E%3Cdefs%3E%3ClinearGradient%20id%3D%22grad1%22%20x1%3D%220%25%22%20y1%3D%220%25%22%20x2%3D%22100%25%22%20y2%3D%22100%25%22%3E%3Cstop%20offset%3D%220%25%22%20style%3D%22stop-color%3A%236d44d9%3Bstop-opacity%3A1%22%20%2F%3E%3Cstop%20offset%3D%22100%25%22%20style%3D%22stop-color%3A%23330B9E%3Bstop-opacity%3A1%22%20%2F%3E%3C%2FlinearGradient%3E%3C%2Fdefs%3E%3Ccircle%20cx%3D%2250%22%20cy%3D%2250%22%20r%3D%2248%22%20fill%3D%22url(%23grad1)%22%20%2F%3E%3Cpath%20fill%3D%22white%22%20d%3D%22M50%2025c-8.28%200-15%206.72-15%2015%200%2011.25%2015%2030%2015%2030s15-18.75%2015-30c0-8.28-6.72-15-15-15zm0%2020.5c-3.04%200-5.5-2.46-5.5-5.5s2.46-5.5%205.5-5.5%205.5%202.46%205.5%205.5-2.46%205.5-5.5%205.5z%22%2F%3E%3C%2Fsvg%3E";

/**
 * CRITICAL: Safe dialog wrapper that validates HTML before showing
 * This prevents "Malformed HTML content" errors
 */
function showModalDialogSafe(
  htmlContent: string,
  title: string,
  width: number,
  height: number,
  context: string
): void {
  try {
    // MANDATORY validation before showing dialog
    if (typeof validateDialogHtml === "function") {
      validateDialogHtml(htmlContent, context);
    } else {
      console.warn("validateDialogHtml function not available - skipping validation");
    }

    // If we get here, HTML is valid - show the dialog
    const output = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(width)
      .setHeight(height);

    SpreadsheetApp.getUi().showModalDialog(output, title);

  } catch (error) {
    // HTML validation failed - show error to user
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error(`Dialog validation failed for ${context}:`, errorMessage);

    // Show simple error dialog instead
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Dialog Error",
      `Unable to show ${context} dialog due to a technical error.\n\nError: ${errorMessage}\n\nPlease report this issue.`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Escapes HTML special characters to prevent XSS attacks
 */
function escapeHtml(unsafe: string): string {
  if (!unsafe) return "";
  return String(unsafe)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

/**
 * Validates and sanitizes HTML message content before passing to generateDialog
 * Use this for messages that contain user-generated or dynamic HTML
 * @param messageLines - Array of message strings that may contain HTML
 * @returns Validated HTML string with proper structure
 */
function validateAndSanitizeMessage(messageLines: string[]): string {
  // Join message lines with proper paragraph tags
  const message = messageLines
    .map((line) => {
      // Check if line already has HTML tags
      if (/<\w+/.test(line)) {
        // Validate the HTML structure
        const validation = typeof validateHtmlContent === "function"
          ? validateHtmlContent(line)
          : { isValid: true, errors: [] };

        if (!validation.isValid) {
          console.warn(`Invalid HTML in message line: ${line}`);
          console.warn(`Errors: ${validation.errors.join(", ")}`);
          // Escape the entire line if it has invalid HTML
          return `<p>${escapeHtml(line)}</p>`;
        }
        return line;
      } else {
        // Plain text - wrap in paragraph tags
        return `<p>${line}</p>`;
      }
    })
    .join("\n");

  return message;
}

/**
 * Validates that a string contains only safe function identifier characters
 * to prevent JavaScript injection
 */
function sanitizeFunctionName(funcName: string): string {
  if (!funcName) return "";
  // Only allow alphanumeric, underscore, and dot (for namespaced functions)
  const sanitized = funcName.replace(/[^a-zA-Z0-9_.]/g, "");
  // Prevent empty or invalid function names
  if (!sanitized || !/^[a-zA-Z_][a-zA-Z0-9_.]*$/.test(sanitized)) {
    console.error(`Invalid function name: ${funcName}`);
    return "";
  }
  return sanitized;
}

function generateDialog(
  title: string,
  message: string,
  buttonText?: string,
  buttonUrl?: string,
  buttonFunction?: string,
  secondaryButtonText?: string,
  secondaryButtonFunction?: string,
): string {
  // Escape user-controlled inputs
  const safeTitle = escapeHtml(title);
  const safeButtonText = escapeHtml(buttonText || "");
  const safeSecondaryButtonText = escapeHtml(secondaryButtonText || "");
  const safeButtonUrl = escapeHtml(buttonUrl || "");

  // Sanitize function names to prevent JavaScript injection
  const safeButtonFunction = sanitizeFunctionName(buttonFunction || "");
  const safeSecondaryButtonFunction = sanitizeFunctionName(secondaryButtonFunction || "");

  // Note: message is not escaped as it contains intentional HTML from calling functions
  const html = `
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
        body {
          font-family: 'Roboto', sans-serif;
          line-height: 1.6;
          color: #e0e0e0;
          background: linear-gradient(135deg, #1a1a1a, #2c2c2c);
          max-width: 800px;
          margin: 0 auto;
          padding: 40px;
          text-align: center;
          box-shadow: 0 0 20px rgba(0, 0, 0, 0.5);
          border-radius: 10px;
        }
        h1 {
          color: #6d44d9;
          font-size: 1.2em;
          text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.3);
          margin-bottom: 30px;
        }
        p {
          margin-bottom: 25px;
          font-size: 1.1em;
        }
        .logo {
          width: 80px;
          height: 80px;
          margin-bottom: 30px;
          border-radius: 50%;
          box-shadow: 0 0 15px rgba(109, 68, 217, 0.7);
          transition: transform 0.3s ease;
        }
        .logo:hover {
          transform: scale(1.05);
        }
        .action-btn {
          display: inline-block;
          background: linear-gradient(45deg, #330B9E, #6d44d9);
          color: white;
          padding: 15px 30px;
          text-decoration: none;
          border-radius: 50px;
          font-weight: bold;
          font-size: 1.2em;
          transition: all 0.3s ease;
          box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
          border: none;
          cursor: pointer;
        }
        .action-btn:hover {
          background: linear-gradient(45deg, #4A0ED6, #8a67e8);
          transform: translateY(-2px);
          box-shadow: 0 6px 8px rgba(0, 0, 0, 0.2);
        }
        .container {
          background-color: rgba(255, 255, 255, 0.05);
          padding: 30px;
          border-radius: 15px;
          margin-top: 30px;
        }
        a {
          font-weight: bold;
          color: #ffffff;
          text-decoration: none;
          position: relative;
          transition: color 0.3s ease;
        }
        a::before {
          content: '';
          position: absolute;
          bottom: 0;
          left: 0;
          width: 0;
          height: 100%;
          background-color: rgba(255, 255, 255, 0.9);
          transition: width 0.3s ease;
          z-index: -1;
        }
        a:hover {
          color: #330B9E;
        }
        a:hover::before {
          width: 100%;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        .spinner {
          display: none;
          width: 20px;
          height: 20px;
          margin: 0 10px;
          border: 3px solid #f3f3f3;
          border-top: 3px solid #6d44d9;
          border-radius: 50%;
          animation: spin 1s linear infinite;
        }
        .processing .spinner {
          display: inline-block;
        }
        .processing .btn-text {
          display: none;
        }
      </style>
    </head>
    <body>
      <img src="${COMAPEO_LOGO_SVG}" alt="CoMapeo Logo" class="logo">
      <h1>${safeTitle}</h1>
      <div class="container">
        ${message}
        ${safeButtonUrl ? `<a href="${safeButtonUrl}" target="_blank" class="action-btn">${safeButtonText}</a>` : ""}
        ${
          safeButtonFunction || safeSecondaryButtonFunction
            ? `
          <div class="button-container" style="display: flex; gap: 15px; justify-content: center; margin-top: 20px;">
            ${safeButtonFunction ? `
              <button onclick="handlePrimaryClick()" class="action-btn primary-btn">
                <span class="btn-text">${safeButtonText}</span>
                <span class="spinner"></span>
              </button>` : ''}
            ${safeSecondaryButtonFunction ? `
              <button onclick="handleSecondaryClick()" class="action-btn secondary-btn" style="background: linear-gradient(45deg, #666, #888) !important;">
                <span class="btn-text">${safeSecondaryButtonText}</span>
                <span class="spinner"></span>
              </button>` : ''}
          </div>
          <script>
            function handlePrimaryClick() {
              const button = document.querySelector('.primary-btn');
              button.classList.add('processing');
              try {
                ${safeButtonFunction}();
              } catch (error) {
                console.error('Error:', error);
                button.classList.remove('processing');
              }
            }
            ${safeSecondaryButtonFunction ? `
            function handleSecondaryClick() {
              const button = document.querySelector('.secondary-btn');
              button.classList.add('processing');
              try {
                ${safeSecondaryButtonFunction}();
              } catch (error) {
                console.error('Error:', error);
                button.classList.remove('processing');
              }
            }` : ''}
          </script>
        `
            : ""
        }
      </div>
    </body>
    </html>
  `;

  // Validate HTML before returning to prevent "Malformed HTML content" errors
  if (typeof validateDialogHtml === "function") {
    try {
      validateDialogHtml(html, title || "Dialog");
    } catch (error) {
      console.error("HTML validation error in generateDialog:", error);
      // Return a safe fallback dialog
      return `
        <!DOCTYPE html>
        <html lang="en">
        <head>
          <meta charset="UTF-8">
          <title>Error</title>
        </head>
        <body>
          <h1>Dialog Error</h1>
          <p>Unable to generate dialog due to malformed HTML content.</p>
          <p>Error: ${escapeHtml(String(error))}</p>
        </body>
        </html>
      `;
    }
  }

  return html;
}

function showIconsGeneratedDialog(folderUrl: string) {
  const title = iconDialogTexts[locale].title;
  const message = validateAndSanitizeMessage(iconDialogTexts[locale].message);
  const buttonText = iconDialogTexts[locale].buttonText;
  const html = generateDialog(title, message, buttonText, folderUrl);
  showModalDialogSafe(html, title, 800, 600, "Icons Generated");
}

function showProcessingModalDialog(dialogText: DialogText) {
  const messageLines = Array.isArray(dialogText.message)
    ? dialogText.message
    : [dialogText.message];

  // Use validateAndSanitizeMessage to ensure proper HTML structure
  const message = validateAndSanitizeMessage(messageLines);
  const html = generateDialog(dialogText.title, message);
  showModalDialogSafe(html, dialogText.title, 800, 600, "Processing");
}

function closeProcessingModalDialog(): void {
  try {
    const autoCloseHtml = HtmlService.createHtmlOutput(
      '<!DOCTYPE html><html><body><script>' +
      'google.script.host.close();' +
      '</script></body></html>'
    )
      .setWidth(10)
      .setHeight(10);
    SpreadsheetApp.getUi().showModalDialog(autoCloseHtml, '');
    Utilities.sleep(50); // Allow dialog to render and close
  } catch (error) {
    console.warn('Unable to close processing dialog:', error);
  }
}

/**
 * Updates the processing dialog with a new progress message without closing it.
 * Uses Google Apps Script's sidebar update mechanism.
 * Note: Due to Apps Script limitations, we can't update modals in real-time.
 * This function will close and reopen the dialog with updated content.
 */
function updateProcessingDialogProgress(mainMessage: string, detailMessage?: string) {
  const title = processingDialogTitle[locale];
  const messageLines = detailMessage
    ? [mainMessage, detailMessage]
    : [mainMessage];

  // Use validateAndSanitizeMessage to ensure proper HTML structure
  const message = validateAndSanitizeMessage(messageLines);
  const html = generateDialog(title, message);

  // Apps Script limitation: We have to close and reopen to update
  // This creates a brief flicker but provides progress feedback
  showModalDialogSafe(html, title, 800, 600, "Processing Progress");
}

function showConfigurationGeneratedDialog(folderUrl: string) {
  const title = generatedConfigDialogTexts[locale].title;
  const message = validateAndSanitizeMessage(generatedConfigDialogTexts[locale].message);
  const buttonText = generatedConfigDialogTexts[locale].buttonText;
  const html = generateDialog(title, message, buttonText, folderUrl);
  showModalDialogSafe(html, title, 800, 980, "Configuration Generated");
}

function showHelpDialog() {
  const title = helpDialogTexts[locale].title;
  const msgHeader = validateAndSanitizeMessage(helpDialogTexts[locale].message);

  // Instructions already contain HTML tags (like <br /> and <a>), validate them
  const instructions = helpDialogTexts[locale].instructions
    .map((instruction) => {
      // Validate each instruction's HTML if it contains tags
      if (/<\w+/.test(instruction)) {
        const validation = typeof validateHtmlContent === "function"
          ? validateHtmlContent(instruction)
          : { isValid: true, errors: [] };

        if (!validation.isValid) {
          console.warn(`Invalid HTML in help instruction: ${instruction}`);
          return `<li>${escapeHtml(instruction)}</li>`;
        }
      }
      return `<li>${instruction}</li>`;
    })
    .join("\n");

  const footer = `<p>${escapeHtml(helpDialogTexts[locale].footer)}</p>`;

  const message = `
  ${msgHeader}
  <ol style="text-align: left;">
  ${instructions}
  </ol>
  ${footer}
`;
  const buttonText = helpDialogTexts[locale].buttonText;
  const buttonUrl =
    "https://github.com/digidem/comapeo-config-spreadsheet-plugin";
  const html = generateDialog(title, message, buttonText, buttonUrl);
  showModalDialogSafe(html, title, 800, 980, "Help");
}

function showSelectTranslationLanguagesDialog() {
  const primaryLanguage = getPrimaryLanguage();
  const availableTargetLanguages = getAvailableTargetLanguages();
  const dialogText = selectTranslationLanguagesDialogText[locale];

  const title = dialogText.title;
  const messageTemplate = dialogText.message;

  const messages = messageTemplate.map((msg) =>
    msg.replace("{{sourceLanguage}}", escapeHtml(primaryLanguage.name)),
  );

  const manualDescription = dialogText.manualSectionDescription
    .map((desc) => `<p>${escapeHtml(desc)}</p>`)
    .join("");

  const languagesJson = JSON.stringify(availableTargetLanguages);
  const primaryLanguageJson = JSON.stringify(primaryLanguage);
  const validationMessagesJson = JSON.stringify(dialogText.validationMessages);
  const manualPlaceholdersJson = JSON.stringify({
    name: dialogText.manualNamePlaceholder,
    iso: dialogText.manualIsoPlaceholder,
  });

  const message = `
    ${messages.map((msg) => `<p>${escapeHtml(msg)}</p>`).join("")}
    <div class="language-selection-container">
      <div class="source-language-info">
        <strong>Source Language:</strong> <span id="sourceLangDisplay"></span>
      </div>

      <div class="search-container">
        <input type="text" id="languageSearch" placeholder="Search languages..." oninput="filterLanguages()" />
      </div>

      <div class="target-languages">
        <strong>Target Languages:</strong>
        <div class="language-count">
          <span id="languageCount">0</span> languages available
        </div>
        <div class="languages-grid" id="languagesGrid"></div>
      </div>

      <div class="select-all-container">
        <button type="button" onclick="selectAllLanguages()" class="select-all-btn">Select All</button>
        <button type="button" onclick="deselectAllLanguages()" class="select-all-btn">Deselect All</button>
        <button type="button" onclick="selectCommonLanguages()" class="select-all-btn">Select Common</button>
      </div>
    </div>

    <div class="custom-language-container">
      <h3>${escapeHtml(dialogText.manualSectionTitle)}</h3>
      ${manualDescription}
      <div class="custom-language-controls">
        <select id="customLanguageSelect" class="language-select" onchange="prefillCustomLanguageRow()">
          <option value="">${escapeHtml(dialogText.manualDropdownPlaceholder)}</option>
        </select>
        <button type="button" onclick="addCustomLanguageRow()" class="select-all-btn custom-add-btn">${escapeHtml(dialogText.manualAddButton)}</button>
      </div>
      <div id="customLanguageRows" class="custom-language-rows">
        <div class="custom-language-row">
          <input type="text" class="custom-language-name" placeholder="${escapeHtml(dialogText.manualNamePlaceholder)}" />
          <input type="text" class="custom-language-iso" placeholder="${escapeHtml(dialogText.manualIsoPlaceholder)}" />
        </div>
      </div>
    </div>

    <style>
      .language-selection-container {
        text-align: left;
        margin: 20px 0;
        padding: 20px;
        background-color: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
      }
      .source-language-info {
        margin-bottom: 20px;
        font-size: 1.1em;
      }
      .search-container {
        margin-bottom: 20px;
      }
      .search-container input {
        width: 100%;
        padding: 10px;
        border: 1px solid #6d44d9;
        border-radius: 5px;
        background-color: rgba(255, 255, 255, 0.1);
        color: #e0e0e0;
        font-size: 1em;
      }
      .search-container input::placeholder {
        color: #a0a0a0;
      }
      .target-languages {
        margin-bottom: 20px;
      }
      .language-count {
        margin: 10px 0;
        font-size: 0.9em;
        color: #a0a0a0;
      }
      .languages-grid {
        max-height: 300px;
        overflow-y: auto;
        border: 1px solid rgba(255, 255, 255, 0.2);
        border-radius: 5px;
        padding: 10px;
        background-color: rgba(255, 255, 255, 0.05);
      }
      .language-checkbox {
        margin: 8px 0;
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .language-checkbox input[type="checkbox"] {
        width: 18px;
        height: 18px;
        margin: 0;
      }
      .language-checkbox label {
        font-size: 1em;
        cursor: pointer;
      }
      .select-all-container {
        display: flex;
        gap: 10px;
        margin-top: 15px;
        flex-wrap: wrap;
      }
      .select-all-btn {
        padding: 8px 16px;
        background-color: rgba(109, 68, 217, 0.3);
        border: 1px solid #6d44d9;
        color: #e0e0e0;
        border-radius: 5px;
        cursor: pointer;
        font-size: 0.9em;
      }
      .select-all-btn:hover {
        background-color: rgba(109, 68, 217, 0.5);
      }
      .custom-language-container {
        margin-top: 30px;
        padding: 20px;
        background-color: rgba(255, 255, 255, 0.08);
        border-radius: 10px;
      }
      .custom-language-controls {
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        margin-bottom: 15px;
      }
      .custom-language-controls select {
        flex: 1 1 260px;
        padding: 10px;
        border: 1px solid #6d44d9;
        border-radius: 5px;
        background-color: rgba(255, 255, 255, 0.1);
        color: #e0e0e0;
      }
      .custom-language-rows {
        display: flex;
        flex-direction: column;
        gap: 10px;
      }
      .custom-language-row {
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
      }
      .custom-language-row input {
        flex: 1 1 200px;
        padding: 8px 12px;
        border: 1px solid rgba(255, 255, 255, 0.2);
        border-radius: 5px;
        background-color: rgba(255, 255, 255, 0.08);
        color: #e0e0e0;
      }
      .custom-language-row input::placeholder {
        color: #a0a0a0;
      }
    </style>
  `;

  const buttonText = dialogText.buttonText;
  const skipButtonText = dialogText.skipButtonText;

  const allJavaScriptLines = [
    "// Data from server (safe JSON)",
    "var availableLanguages = " + languagesJson + ";",
    "var primaryLanguage = " + primaryLanguageJson + ";",
    "var validationMessages = " + validationMessagesJson + ";",
    "var manualPlaceholders = " + manualPlaceholdersJson + ";",
    "var customIsoPattern = /^[a-z]{2,8}(-[a-z]{2,8})?$/;",
    "var commonLanguages = ['en', 'es', 'pt', 'fr', 'de', 'it', 'ja', 'ko', 'zh-CN', 'ru', 'ar', 'hi'];",
    "",
    "document.addEventListener('DOMContentLoaded', function() {",
    "  initializeLanguageSelection();",
    "  populateCustomLanguageOptions();",
    "});",
    "",
    "function initializeLanguageSelection() {",
    "  document.getElementById('sourceLangDisplay').textContent =",
    "    primaryLanguage.name + ' (' + primaryLanguage.code + ')';",
    "",
    "  renderLanguageCheckboxes(availableLanguages);",
    "}",
    "",
    "function renderLanguageCheckboxes(languages) {",
    "  var sortedLanguages = Object.entries(languages).sort(function(a, b) {",
    "    return a[1].localeCompare(b[1]);",
    "  });",
    "",
    "  var grid = document.getElementById('languagesGrid');",
    "  grid.innerHTML = '';",
    "",
    "  for (var i = 0; i < sortedLanguages.length; i++) {",
    "    var code = sortedLanguages[i][0];",
    "    var name = sortedLanguages[i][1];",
    "",
    "    var checkbox = document.createElement('div');",
    "    checkbox.className = 'language-checkbox';",
    "",
    "    var input = document.createElement('input');",
    "    input.type = 'checkbox';",
    "    input.id = 'lang_' + code;",
    "    input.value = code;",
    "",
    "    var label = document.createElement('label');",
    "    label.htmlFor = 'lang_' + code;",
    "    label.textContent = name + ' (' + code + ')';",
    "",
    "    checkbox.appendChild(input);",
    "    checkbox.appendChild(label);",
    "    grid.appendChild(checkbox);",
    "  }",
    "",
    "  document.getElementById('languageCount').textContent = sortedLanguages.length;",
    "}",
    "",
    "function filterLanguages() {",
    "  var searchTerm = document.getElementById('languageSearch').value.toLowerCase();",
    "  var checkboxes = document.querySelectorAll('.language-checkbox');",
    "  var visibleCount = 0;",
    "",
    "  for (var i = 0; i < checkboxes.length; i++) {",
    "    var checkbox = checkboxes[i];",
    "    var label = checkbox.querySelector('label').textContent.toLowerCase();",
    "    if (label.includes(searchTerm)) {",
    "      checkbox.style.display = 'flex';",
    "      visibleCount++;",
    "    } else {",
    "      checkbox.style.display = 'none';",
    "    }",
    "  }",
    "",
    "  document.getElementById('languageCount').textContent = visibleCount;",
    "}",
    "",
    "function selectAllLanguages() {",
    "  var allCheckboxes = document.querySelectorAll('input[type=\"checkbox\"]');",
    "  for (var i = 0; i < allCheckboxes.length; i++) {",
    "    var cb = allCheckboxes[i];",
    "    var parent = cb.closest('.language-checkbox');",
    "    if (parent && parent.style.display !== 'none') {",
    "      cb.checked = true;",
    "    }",
    "  }",
    "}",
    "",
    "function deselectAllLanguages() {",
    "  var checkboxes = document.querySelectorAll('input[type=\"checkbox\"]');",
    "  for (var i = 0; i < checkboxes.length; i++) {",
    "    checkboxes[i].checked = false;",
    "  }",
    "}",
    "",
    "function selectCommonLanguages() {",
    "  deselectAllLanguages();",
    "  for (var i = 0; i < commonLanguages.length; i++) {",
    "    var langCode = commonLanguages[i];",
    "    var checkbox = document.getElementById('lang_' + langCode);",
    "    if (checkbox) {",
    "      checkbox.checked = true;",
    "    }",
    "  }",
    "}",
    "",
    "function createCustomLanguageRow() {",
    "  var row = document.createElement('div');",
    "  row.className = 'custom-language-row';",
    "  var nameInput = document.createElement('input');",
    "  nameInput.type = 'text';",
    "  nameInput.className = 'custom-language-name';",
    "  nameInput.placeholder = manualPlaceholders.name;",
    "  var isoInput = document.createElement('input');",
    "  isoInput.type = 'text';",
    "  isoInput.className = 'custom-language-iso';",
    "  isoInput.placeholder = manualPlaceholders.iso;",
    "  row.appendChild(nameInput);",
    "  row.appendChild(isoInput);",
    "  return row;",
    "}",
    "",
    "function addCustomLanguageRow() {",
    "  var container = document.getElementById('customLanguageRows');",
    "  if (!container) {",
    "    return;",
    "  }",
    "  container.appendChild(createCustomLanguageRow());",
    "}",
    "",
    "function populateCustomLanguageOptions() {",
    "  var select = document.getElementById('customLanguageSelect');",
    "  if (!select) {",
    "    return;",
    "  }",
    "  var sortedLanguages = Object.entries(availableLanguages).sort(function(a, b) {",
    "    return a[1].localeCompare(b[1]);",
    "  });",
    "  for (var i = 0; i < sortedLanguages.length; i++) {",
    "    var option = document.createElement('option');",
    "    option.value = sortedLanguages[i][0];",
    "    option.textContent = sortedLanguages[i][1] + ' (' + sortedLanguages[i][0] + ')';",
    "    select.appendChild(option);",
    "  }",
    "}",
    "",
    "function prefillCustomLanguageRow() {",
    "  var select = document.getElementById('customLanguageSelect');",
    "  if (!select || !select.value) {",
    "    return;",
    "  }",
    "  var rows = document.querySelectorAll('.custom-language-row');",
    "  if (rows.length === 0) {",
    "    addCustomLanguageRow();",
    "  }",
    "  var targetRow = rows[rows.length - 1];",
    "  var option = select.options[select.selectedIndex];",
    "  var label = option.text;",
    "  var name = label.split(' (')[0];",
    "  var iso = option.value;",
    "  var nameInput = targetRow.querySelector('.custom-language-name');",
    "  var isoInput = targetRow.querySelector('.custom-language-iso');",
    "  if (nameInput) { nameInput.value = name; }",
    "  if (isoInput) { isoInput.value = iso; }",
    "  select.value = '';",
    "}",
    "",
    "function collectCustomLanguages() {",
    "  var rows = document.querySelectorAll('.custom-language-row');",
    "  var customLanguages = [];",
    "  var seenIso = {};",
    "",
    "  for (var i = 0; i < rows.length; i++) {",
    "    var row = rows[i];",
    "    var nameInput = row.querySelector('.custom-language-name');",
    "    var isoInput = row.querySelector('.custom-language-iso');",
    "    var name = nameInput ? nameInput.value.trim() : '';",
    "    var iso = isoInput ? isoInput.value.trim() : '';",
    "",
    "    if (name === '' && iso === '') {",
    "      continue;",
    "    }",
    "",
    "    if (name === '' || iso === '') {",
    "      alert(validationMessages.missingCustomFields);",
    "      return null;",
    "    }",
    "",
    "    var normalizedIso = iso.toLowerCase();",
    "    if (!customIsoPattern.test(normalizedIso)) {",
    "      alert(validationMessages.invalidCustomIso);",
    "      return null;",
    "    }",
    "",
    "    if (seenIso[normalizedIso]) {",
    "      alert(validationMessages.duplicateCustomIso);",
    "      return null;",
    "    }",
    "",
    "    seenIso[normalizedIso] = true;",
    "    customLanguages.push({ name: name, iso: normalizedIso });",
    "  }",
    "",
    "  return customLanguages;",
    "}",
    "",
    "function resetProcessingState(selector) {",
    "  var button = document.querySelector(selector);",
    "  if (button) {",
    "    button.classList.remove('processing');",
    "  }",
    "}",
    "",
    "function getSelectedTargetLanguages() {",
    "  var checkedBoxes = document.querySelectorAll('input[type=\"checkbox\"]:checked');",
    "  var selectedLanguages = [];",
    "  for (var i = 0; i < checkedBoxes.length; i++) {",
    "    selectedLanguages.push(checkedBoxes[i].value);",
    "  }",
    "",
    "  var customLanguages = collectCustomLanguages();",
    "  if (customLanguages === null) {",
    "    resetProcessingState('.primary-btn');",
    "    return;",
    "  }",
    "",
    "  if (selectedLanguages.length === 0) {",
    "    alert(validationMessages.noAutoSelection);",
    "    resetProcessingState('.primary-btn');",
    "    return;",
    "  }",
    "",
    "  google.script.run",
    "    .withFailureHandler(function(error) {",
    "      console.error('Failed to process language selection:', error);",
    "      resetProcessingState('.primary-btn');",
    "      alert('Operation failed: ' + error.message);",
    "    })",
    "    .withSuccessHandler(function() {",
    "      google.script.host.close();",
    "    })",
    "    .handleLanguageSelection({",
    "      autoTranslateLanguages: selectedLanguages,",
    "      customLanguages: customLanguages",
    "    });",
    "}",
    "",
    "function cancelTranslationSelection() {",
    "  console.log('[CLIENT] Cancel Translation dialog');",
    "  resetProcessingState('.secondary-btn');",
    "  google.script.host.close();",
    "}",
    "",
    "window.filterLanguages = filterLanguages;",
    "window.selectAllLanguages = selectAllLanguages;",
    "window.deselectAllLanguages = deselectAllLanguages;",
    "window.selectCommonLanguages = selectCommonLanguages;",
    "window.addCustomLanguageRow = addCustomLanguageRow;",
    "window.prefillCustomLanguageRow = prefillCustomLanguageRow;",
    "window.getSelectedTargetLanguages = getSelectedTargetLanguages;",
    "window.cancelTranslationSelection = cancelTranslationSelection;"
  ];
  const allJavaScript = allJavaScriptLines.join("\n");

  const fullMessage = message + '<script>' + allJavaScript + '</script>';

  const html = generateDialog(
    title,
    fullMessage,
    buttonText,
    null,
    "getSelectedTargetLanguages",
    skipButtonText,
    "cancelTranslationSelection",
  );
  showModalDialogSafe(html, title, 820, 950, "Translation Language Selection");
}

/**
 * Get user-friendly label for error type
 */
function getErrorTypeLabel(errorType: string): string {
  const labels: Record<string, string> = {
    format: "📄 Format Errors",
    permission: "🔒 Permission Errors",
    api: "🌐 API Errors",
    drive: "💾 Drive Errors",
    validation: "✓ Validation Errors",
    network: "📡 Network Errors",
    timeout: "⏱️ Timeout Errors",
    unknown: "❓ Unknown Errors",
  };
  return labels[errorType] || errorType;
}

/**
 * Show icon processing error report dialog
 * @param errorSummary - Summary of icon errors from processing
 */
function showIconErrorDialog(errorSummary: IconErrorSummary): void {
  const title = iconErrorDialogTexts[locale].title;

  // Build summary message
  const summaryLines: string[] = [];
  if (errorSummary.errorCount === 0) {
    summaryLines.push(
      `<p class="success">✅ All ${errorSummary.totalProcessed} icons processed successfully</p>`,
    );
  } else {
    summaryLines.push(
      `<p>Processed <strong>${errorSummary.totalProcessed}</strong> icon(s):</p>`,
    );
    summaryLines.push(`<ul>`);
    summaryLines.push(
      `  <li>✅ <strong>${errorSummary.successCount}</strong> successful</li>`,
    );
    summaryLines.push(
      `  <li>⚠️ <strong>${errorSummary.errorCount}</strong> with issues</li>`,
    );
    if (errorSummary.fallbackCount > 0) {
      summaryLines.push(
        `  <li>🔄 <strong>${errorSummary.fallbackCount}</strong> using fallback icon</li>`,
      );
    }
    summaryLines.push(`</ul>`);
  }

  // Group errors by type
  const errorsByType = errorSummary.errorsByType;
  const errorSections: string[] = [];

  if (errorsByType.size > 0) {
    errorSections.push(`<h3>Issues by Category:</h3>`);

    errorsByType.forEach((errors, errorType) => {
      errorSections.push(`<div class="error-category">`);
      errorSections.push(
        `  <h4>${getErrorTypeLabel(errorType)} (${errors.length})</h4>`,
      );
      errorSections.push(`  <ul class="error-list">`);

      // Show first 5 errors, collapse rest
      const visibleErrors = errors.slice(0, 5);
      visibleErrors.forEach((error) => {
        errorSections.push(`    <li>`);
        errorSections.push(
          `      <strong>${escapeHtml(error.iconName)}</strong>: ${escapeHtml(error.userMessage)}`,
        );
        errorSections.push(
          `      <br/><span class="suggestion">→ ${escapeHtml(error.suggestedAction)}</span>`,
        );
        if (error.context && error.context.fileId) {
          errorSections.push(
            `      <br/><span class="context">File ID: ${escapeHtml(error.context.fileId)}</span>`,
          );
        }
        errorSections.push(`    </li>`);
      });

      if (errors.length > 5) {
        errorSections.push(
          `    <li class="more-errors">... and ${errors.length - 5} more</li>`,
        );
      }

      errorSections.push(`  </ul>`);
      errorSections.push(`</div>`);
    });
  }

  const message = `
    <div class="icon-error-report">
      ${summaryLines.join("\n")}
      ${errorSections.join("\n")}
      ${errorSummary.hasCriticalErrors
        ? `<div class="critical-warning">⚠️ Critical errors detected. Please address permission and validation issues.</div>`
        : ""}
    </div>
    <style>
      .icon-error-report { text-align: left; }
      .success { color: #4caf50; font-size: 1.1em; }
      .error-category { margin: 15px 0; padding: 10px; background: rgba(255,255,255,0.05); border-radius: 5px; }
      .error-category h4 { margin: 5px 0; color: #ff9800; }
      .error-list { list-style: none; padding-left: 10px; }
      .error-list li { margin: 8px 0; padding: 5px; background: rgba(255,255,255,0.02); border-left: 3px solid #ff9800; }
      .suggestion { color: #81c784; font-size: 0.9em; }
      .context { color: #90caf9; font-size: 0.85em; }
      .more-errors { font-style: italic; color: #999; }
      .critical-warning { margin-top: 20px; padding: 15px; background: rgba(255,152,0,0.2); border: 2px solid #ff9800; border-radius: 5px; font-weight: bold; }
    </style>
  `;

  const html = generateDialog(
    title,
    message,
    iconErrorDialogTexts[locale].downloadButtonText,
    undefined,
    "downloadIconErrorReport",
    errorSummary.errorCount > 0
      ? iconErrorDialogTexts[locale].continueButtonText
      : iconErrorDialogTexts[locale].okButtonText,
    "google.script.host.close",
  );

  showModalDialogSafe(html, title, 800, 600, "Icon Error Report");
}

/**
 * Download icon error report as CSV
 * Called from the error dialog
 */
function downloadIconErrorReport(): void {
  // This function would be called from the dialog
  // It should retrieve the error summary and trigger CSV download
  console.log(
    "Download icon error report - implementation depends on error storage strategy",
  );
  // Note: Full implementation requires error summary to be stored in PropertiesService
  // and retrieved here, then converted to CSV using collector.toCSV()
}
