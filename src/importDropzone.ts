// External function declarations
declare function createOrClearSheet(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: string,
): GoogleAppsScript.Spreadsheet.Sheet;
declare function processImportedCategoryFile(
  fileName: string,
  base64Data: string,
): { success: boolean; message: string; details?: any; warnings?: string[] };
declare function extractConfigurationData(
  extractedFiles: GoogleAppsScript.Base.Blob[],
  tempFolder: GoogleAppsScript.Drive.Folder,
): any;
declare function normalizeConfig(jsonData: any): any;
declare function applyConfigurationToSpreadsheet(config: any): void;

/**
 * Interface for dropzone configuration options
 */
interface DropzoneOptions {
  acceptedFileTypes: string[];
  maxFileSize: number; // in bytes
  multiple: boolean;
}

/**
 * Shows the import dialog with a dropzone for file uploads
 */
function showImportDropzoneDialog(): void {
  const htmlOutput = HtmlService.createHtmlOutput(createDropzoneHtml())
    .setWidth(600)
    .setHeight(400)
    .setTitle("Import Category File");

  SpreadsheetApp.getUi().showModalDialog(
    htmlOutput,
    "Import Category File",
  );
}

/**
 * Creates the HTML for the dropzone dialog
 */
function createDropzoneHtml(): string {
  const dropzoneOptions: DropzoneOptions = {
    acceptedFileTypes: [".comapeocat", ".mapeosettings"],
    maxFileSize: 50 * 1024 * 1024, // 50MB
    multiple: false,
  };

  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <title>Import Category File</title>
  <style>
    body {
      font-family: 'Roboto', Arial, sans-serif;
      margin: 0;
      padding: 20px;
      color: #333;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 500px;
      margin: 0 auto;
      background: white;
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
      padding: 20px;
    }
    h1 {
      font-size: 1.5em;
      margin-top: 0;
      margin-bottom: 20px;
      color: #4285f4;
      text-align: center;
    }
    .dropzone {
      border: 2px dashed #4285f4;
      border-radius: 5px;
      padding: 25px;
      text-align: center;
      cursor: pointer;
      transition: all 0.3s ease;
      background-color: #f8f9fa;
      position: relative;
      min-height: 150px;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
    }
    .dropzone.dragover {
      background-color: #e8f0fe;
      border-color: #1a73e8;
    }
    .dropzone.error {
      border-color: #ea4335;
      background-color: #fce8e6;
    }
    .dropzone.success {
      border-color: #34a853;
      background-color: #e6f4ea;
    }
    .dropzone-icon {
      font-size: 48px;
      margin-bottom: 10px;
      color: #4285f4;
    }
    .dropzone-text {
      margin-bottom: 10px;
      font-size: 16px;
    }
    .dropzone-subtext {
      font-size: 12px;
      color: #5f6368;
    }
    .file-input {
      display: none;
    }
    .progress-container {
      width: 100%;
      margin-top: 15px;
      display: none;
    }
    .progress-bar {
      height: 8px;
      background-color: #e0e0e0;
      border-radius: 4px;
      overflow: hidden;
      position: relative;
    }
    .progress {
      height: 100%;
      background-color: #4285f4;
      width: 0%;
      transition: width 0.3s ease;
      position: relative;
    }

    .progress::after {
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: linear-gradient(
        -45deg,
        rgba(255, 255, 255, 0.2) 25%,
        transparent 25%,
        transparent 50%,
        rgba(255, 255, 255, 0.2) 50%,
        rgba(255, 255, 255, 0.2) 75%,
        transparent 75%
      );
      background-size: 16px 16px;
      animation: progress-animation 1s linear infinite;
      z-index: 1;
    }

    @keyframes progress-animation {
      0% {
        background-position: 0 0;
      }
      100% {
        background-position: 16px 0;
      }
    }
    .status-message {
      margin-top: 10px;
      font-size: 14px;
      text-align: center;
      min-height: 20px;
    }
    .error-message {
      color: #ea4335;
      margin-bottom: 15px;
    }
    .success-message {
      color: #34a853;
    }
    .details-text {
      font-size: 12px;
      color: #5f6368;
      display: block;
      margin-top: 5px;
    }
    .error-details {
      margin-top: 10px;
      font-size: 13px;
      text-align: left;
      background-color: #fce8e6;
      border-radius: 4px;
      padding: 10px;
    }
    .error-details ul {
      margin: 5px 0;
      padding-left: 20px;
    }
    .error-stage {
      font-style: italic;
      margin-top: 8px;
      font-size: 12px;
    }
    .warnings-container {
      margin-top: 10px;
      font-size: 13px;
      text-align: left;
      background-color: #fff3e0;
      border-radius: 4px;
      padding: 10px;
      color: #e65100;
    }
    .warnings-container ul {
      margin: 5px 0;
      padding-left: 20px;
    }
    .retry-button, .view-button, .cancel-button {
      display: inline-block;
      margin: 15px 5px 0;
      padding: 8px 16px;
      color: white;
      border: none;
      border-radius: 4px;
      cursor: pointer;
      font-size: 14px;
      transition: all 0.3s;
      box-shadow: 0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24);
    }
    .retry-button {
      background-color: #4285f4;
    }
    .view-button {
      background-color: #34a853;
    }
    .cancel-button {
      background-color: #ea4335;
    }
    .retry-button:hover, .view-button:hover, .cancel-button:hover {
      box-shadow: 0 3px 6px rgba(0,0,0,0.16), 0 3px 6px rgba(0,0,0,0.23);
      transform: translateY(-1px);
    }
    .retry-button:hover {
      background-color: #1a73e8;
    }
    .view-button:hover {
      background-color: #2d9348;
    }
    .cancel-button:hover {
      background-color: #d33426;
    }
    .warning-message {
      background-color: #fff3e0;
      border-left: 4px solid #ff9800;
      padding: 12px;
      margin-bottom: 15px;
      border-radius: 4px;
      font-size: 14px;
      line-height: 1.4;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>Import Category File</h1>
    <div class="warning-message">
      <strong>Warning:</strong> Importing a category file will erase all current spreadsheet data. This action cannot be undone.
    </div>
    <div id="dropzone" class="dropzone">
      <div class="dropzone-icon">📁</div>
      <div class="dropzone-text">Drag and drop your file here</div>
      <div class="dropzone-subtext">or click to browse files</div>
      <div class="dropzone-subtext">Accepted file types: .comapeocat, .mapeosettings</div>
      <input type="file" id="file-input" class="file-input" accept=".comapeocat,.mapeosettings">
      <div class="progress-container" id="progress-container">
        <div class="progress-bar">
          <div class="progress" id="progress"></div>
        </div>
      </div>
    </div>
    <div id="status-message" class="status-message"></div>
  </div>

  <script>
    // @ts-nocheck - This is client-side JavaScript code that will be embedded in HTML
    // Store dropzone options
    const dropzoneOptions = ${JSON.stringify(dropzoneOptions)};

    // DOM elements
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('file-input');
    const progressContainer = document.getElementById('progress-container');
    const progress = document.getElementById('progress');
    const statusMessage = document.getElementById('status-message');

    // Add event listeners
    dropzone.addEventListener('click', () => fileInput.click());
    dropzone.addEventListener('dragover', handleDragOver);
    dropzone.addEventListener('dragleave', handleDragLeave);
    dropzone.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);

    // Handle drag over event
    function handleDragOver(e) {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.add('dragover');
    }

    // Handle drag leave event
    function handleDragLeave(e) {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');
    }

    // Handle drop event
    function handleDrop(e) {
      e.preventDefault();
      e.stopPropagation();
      dropzone.classList.remove('dragover');

      const files = e.dataTransfer.files;
      if (files.length > 0) {
        processFile(files[0]);
      }
    }

    // Handle file selection
    function handleFileSelect(e) {
      const files = e.target.files;
      if (files.length > 0) {
        processFile(files[0]);
      }
    }

    // Process the selected file
    function processFile(file) {
      // Reset UI
      resetUI();

      // Validate file type
      const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
      if (!dropzoneOptions.acceptedFileTypes.includes(fileExtension)) {
        showError('Invalid file type. Please upload a .comapeocat or .mapeosettings file.');
        return;
      }

      // Validate file size
      if (file.size > dropzoneOptions.maxFileSize) {
        showError('File is too large. Maximum file size is ' +
          (dropzoneOptions.maxFileSize / (1024 * 1024)) + 'MB.');
        return;
      }

      // Show progress
      progressContainer.style.display = 'block';
      updateProgress(10);
      statusMessage.textContent = 'Reading file...';

      // Read the file
      const reader = new FileReader();

      reader.onload = function(e) {
        updateProgress(20);
        statusMessage.innerHTML = 'Processing file...<br><span class="details-text">Preparing to extract file contents</span>';

        try {
          // Make sure we're getting the base64 data correctly
          const result = e.target.result.toString();
          const base64data = result.split(',')[1];

          if (!base64data) {
            throw new Error('Failed to extract base64 data from file');
          }

          // Call the appropriate Google Apps Script function based on file extension
          const fileExtension = file.name.toString().split('.').pop().toLowerCase();

          // Create a simple progress handler function
          const progressHandler = function(progressData) {
            try {
              if (progressData && typeof progressData === 'object') {
                // Update progress bar
                if (typeof progressData.percent === 'number') {
                  updateProgress(20 + Math.round(progressData.percent * 0.8)); // Scale to 20-100%
                }

                // Update status message
                if (progressData.stage) {
                  let statusHtml = 'Processing: ' + progressData.stage;

                  // Add details if available
                  if (progressData.detail) {
                    statusHtml += '<br><span class="details-text">' + progressData.detail + '</span>';
                  }

                  statusMessage.innerHTML = statusHtml;
                }
              }
            } catch (e) {
              console.error('Error in progress handler:', e);
            }
          };

          if (fileExtension === 'comapeocat' || fileExtension === 'zip' || fileExtension === 'mapeosettings') {
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .withUserObject({ fileName: file.name })
              .importConfigurationFile(file.name, base64data);
          } else {
            onFailure(new Error('Unsupported file type'));
          }
        } catch (error) {
          onFailure(error);
        }
      };

      reader.onerror = function() {
        onFailure(new Error('Error reading file'));
      };

      reader.readAsDataURL(file);
    }

    // Handle successful import
    function onSuccess(result, userObject) {
      updateProgress(100);
      dropzone.classList.add('success');

      // Display success message with details if available
      let successMessage = 'File imported successfully!';
      if (userObject && userObject.fileName) {
        successMessage = 'File "' + userObject.fileName + '" imported successfully!';
      }

      if (result && result.details) {
        const details = result.details;
        const detailsText = [];

        if (details.presets) detailsText.push(details.presets + ' categories');
        if (details.fields) detailsText.push(details.fields + ' details');
        if (details.icons) detailsText.push(details.icons + ' icons');
        if (details.languages) {
          const langCount = Array.isArray(details.languages) ?
            details.languages.length :
            (typeof details.languages === 'object' ? Object.keys(details.languages).length : 0);

          if (langCount > 0) {
            detailsText.push(langCount + ' languages');
          }
        }

        if (detailsText.length > 0) {
          successMessage += '<br><span class="details-text">Imported: ' + detailsText.join(', ') + '</span>';
        }

        // Add processing time if available
        if (details.processingTime) {
          const seconds = (details.processingTime / 1000).toFixed(2);
          successMessage += '<br><span class="details-text">Processing time: ' + seconds + ' seconds</span>';
        }
      }

      statusMessage.innerHTML = successMessage;
      statusMessage.classList.add('success-message');

      // Show warnings if any
      if (result && result.warnings && result.warnings.length > 0) {
        const warningsContainer = document.createElement('div');
        warningsContainer.className = 'warnings-container';
        warningsContainer.innerHTML = '<strong>Warnings:</strong><ul>' +
          result.warnings.map(function(w) { return '<li>' + (w.userMessage || w) + '</li>'; }).join('') +
          '</ul>';
        statusMessage.appendChild(warningsContainer);
      }

      // Add a "View Spreadsheet" button
      const viewButton = document.createElement('button');
      viewButton.className = 'view-button';
      viewButton.textContent = 'View Spreadsheet';
      viewButton.onclick = function() {
        google.script.host.close();
      };
      statusMessage.appendChild(viewButton);

      // Close the dialog automatically after a longer delay
      setTimeout(() => {
        google.script.host.close();
      }, 5000);
    }

    // Handle import failure
    function onFailure(error, userObject) {
      dropzone.classList.add('error');
      updateProgress(0);

      // Create detailed error message
      let errorMessage = 'Error importing file';
      if (userObject && userObject.fileName) {
        errorMessage = 'Error importing file "' + userObject.fileName + '"';
      }

      if (error) {
        if (typeof error === 'string') {
          errorMessage += ': ' + error;
        } else if (error.message) {
          errorMessage += ': ' + error.message;
        }
      }

      statusMessage.innerHTML = errorMessage;
      statusMessage.classList.add('error-message');

      // Add more details if available
      if (error && error.details) {
        const errorDetails = document.createElement('div');
        errorDetails.className = 'error-details';

        // If there are specific errors in the details
        if (error.details.errors && error.details.errors.length > 0) {
          const errorList = document.createElement('ul');
          error.details.errors.forEach(err => {
            const li = document.createElement('li');
            li.textContent = err.userMessage || err;
            errorList.appendChild(li);
          });
          errorDetails.appendChild(errorList);
        }

        // If there's a stage indication
        if (error.details.stage) {
          const stagePara = document.createElement('p');
          stagePara.className = 'error-stage';
          stagePara.textContent = 'Error occurred during: ' + error.details.stage;
          errorDetails.appendChild(stagePara);
        }

        statusMessage.appendChild(errorDetails);
      }

      // Add a retry button
      const retryButton = document.createElement('button');
      retryButton.className = 'retry-button';
      retryButton.textContent = 'Try Again';
      retryButton.onclick = function() {
        resetUI();
      };
      statusMessage.appendChild(retryButton);

      // Add a cancel button
      const cancelButton = document.createElement('button');
      cancelButton.className = 'cancel-button';
      cancelButton.textContent = 'Cancel';
      cancelButton.onclick = function() {
        google.script.host.close();
      };
      statusMessage.appendChild(cancelButton);
    }

    // Show error message
    function showError(message) {
      dropzone.classList.add('error');
      statusMessage.textContent = message;
      statusMessage.classList.add('error-message');
    }

    // Update progress bar
    function updateProgress(percent) {
      progress.style.width = percent + '%';
    }

    // Reset UI
    function resetUI() {
      dropzone.classList.remove('error', 'success', 'dragover');
      statusMessage.textContent = '';
      statusMessage.classList.remove('error-message', 'success-message');
      updateProgress(0);
    }
  </script>
</body>
</html>`;
}
/**
 * Handles the file import process
 * @param fileName - Name of the uploaded file
 * @param base64Data - Base64 encoded file data
 * @returns Result of the import operation
 */
/**
 * Processes a .mapeosettings file (classic Mapeo configuration)
 * @param fileName - Name of the uploaded file
 * @param base64Data - Base64 encoded file data
 * @returns Result of the import operation
 */
function processMapeoSettingsFile(
  fileName: string,
  base64Data: string,
): { success: boolean; message: string; details?: any } {
  try {
    console.log(`Starting import of Mapeo settings file: ${fileName}`);

    // Decode the base64 data
    console.log("Decoding base64 data...");
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/octet-stream",
      fileName,
    );
    console.log(`Decoded file size: ${blob.getBytes().length} bytes`);

    // Create a temporary folder
    console.log("Creating temporary folder...");
    const tempFolder = DriveApp.createFolder(
      "Mapeo_Settings_Import_" + new Date().getTime(),
    );
    console.log(`Created temporary folder: ${tempFolder.getName()}`);

    // Extract the file - using the direct approach from testImportCategory.ts
    console.log("Extracting file...");
    let extractionResult: {
      success: boolean;
      message: string;
      files?: GoogleAppsScript.Base.Blob[];
      tempFolder?: GoogleAppsScript.Drive.Folder;
      validationErrors?: string[];
      validationWarnings?: string[];
    };

    try {
      // Use the simple call without progress handler
      extractionResult = extractAndValidateFile(fileName, blob);

      if (!extractionResult.success) {
        console.error("Extraction failed:", extractionResult.message);
        return {
          success: false,
          message: extractionResult.message,
        };
      }

      console.log(
        `Extraction successful: ${extractionResult.files.length} files extracted`,
      );
    } catch (error) {
      console.error("Error during extraction:", error);
      return {
        success: false,
        message: `Failed to extract file: ${error instanceof Error ? error.message : String(error)}`,
      };
    }

    // Extract configuration data
    console.log("Extracting configuration data...");
    let configData: {
      metadata: any;
      presets: any[];
      fields: any[];
      icons: any[];
      messages: Record<string, any>;
    };

    try {
      // Use the simple call without progress handler
      configData = extractConfigurationData(
        extractionResult.files,
        extractionResult.tempFolder,
      );

      console.log("Configuration data extracted successfully");
      console.log(
        `- Presets: ${configData.presets ? configData.presets.length : 0}`,
      );
      console.log(
        `- Fields: ${configData.fields ? configData.fields.length : 0}`,
      );
      console.log(`- Icons: ${configData.icons ? configData.icons.length : 0}`);
    } catch (error) {
      console.error("Error extracting configuration data:", error);
      return {
        success: false,
        message: `Failed to extract configuration data: ${error instanceof Error ? error.message : String(error)}`,
      };
    }

    // Apply configuration to spreadsheet
    console.log("Applying configuration to spreadsheet...");
    try {
      applyConfigurationToSpreadsheet(configData);
      console.log("Configuration applied to spreadsheet successfully");
    } catch (error) {
      console.error("Error applying configuration to spreadsheet:", error);
      return {
        success: false,
        message: `Failed to apply configuration to spreadsheet: ${error instanceof Error ? error.message : String(error)}`,
      };
    }

    // Clean up
    console.log("Cleaning up...");
    try {
      extractionResult.tempFolder.setTrashed(true);
      console.log("Temporary folder deleted");
    } catch (error) {
      console.warn("Error cleaning up temporary folder:", error);
    }

    // Return success
    console.log("Import completed successfully");
    return {
      success: true,
      message: "Mapeo settings file imported successfully",
      details: {
        presets: configData.presets ? configData.presets.length : 0,
        fields: configData.fields ? configData.fields.length : 0,
        icons: configData.icons ? configData.icons.length : 0,
      },
    };
  } catch (error) {
    console.error("Error processing Mapeo settings file:", error);
    return {
      success: false,
      message: error instanceof Error ? error.message : String(error),
    };
  }
}

/**
 * Handles the file import process
 * @param fileName - Name of the uploaded file
 * @param base64Data - Base64 encoded file data
 * @returns Result of the import operation
 */
function handleFileImport(
  fileName: string,
  base64Data: string,
): { success: boolean; message: string; details?: any } {
  try {
    // Validate file extension
    const fileExtension = fileName.toString().split(".").pop()?.toLowerCase();
    if (fileExtension !== "comapeocat" && fileExtension !== "mapeosettings") {
      throw new Error(
        `Invalid file type. Please upload a .comapeocat or .mapeosettings file.`,
      );
    }

    // Use the direct import approach for both file types
    return importConfigurationFile(fileName, base64Data);
  } catch (error) {
    console.error("Error importing file:", error);
    return {
      success: false,
      message:
        error instanceof Error ? error.message : "Unknown error occurred",
    };
  }
}

/**
 * Unified function to import configuration files (both .comapeocat and .mapeosettings)
 * This uses the direct approach from testImportCategory.ts
 * @param fileName - Name of the uploaded file
 * @param base64Data - Base64 encoded file data
 * @returns Result of the import operation
 */
function importConfigurationFile(
  fileName: string,
  base64Data: string,
): { success: boolean; message: string; details?: any; warnings?: string[] } {
  try {
    console.log(`Starting import of file: ${fileName}`);

    // Decode the base64 data
    console.log("Decoding base64 data...");
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      "application/octet-stream",
      fileName,
    );
    console.log(`Decoded file size: ${blob.getBytes().length} bytes`);

    // Create a temporary folder
    console.log("Creating temporary folder...");
    const tempFolder = DriveApp.createFolder(
      "Config_Import_" + new Date().getTime(),
    );
    console.log(`Created temporary folder: ${tempFolder.getName()}`);

    // Extract the file - using the direct approach from testImportCategory.ts
    console.log("Extracting file...");
    let extractionResult: {
      success: boolean;
      message: string;
      files?: GoogleAppsScript.Base.Blob[];
      tempFolder?: GoogleAppsScript.Drive.Folder;
      validationErrors?: string[];
      validationWarnings?: string[];
    };

    try {
      // Use the simple call without progress handler
      extractionResult = extractAndValidateFile(fileName, blob);

      if (!extractionResult.success) {
        console.error("Extraction failed:", extractionResult.message);
        return {
          success: false,
          message: extractionResult.message,
        };
      }

      console.log(
        `Extraction successful: ${extractionResult.files.length} files extracted`,
      );
    } catch (error) {
      console.error("Error during extraction:", error);
      return {
        success: false,
        message: `Failed to extract file: ${error instanceof Error ? error.message : String(error)}`,
      };
    }

    // Extract configuration data
    console.log("Extracting configuration data...");
    let configData: {
      metadata: any;
      presets: any[];
      fields: any[];
      icons: any[];
      messages: Record<string, any>;
    };

    try {
      // Use the simple call without progress handler
      configData = extractConfigurationData(
        extractionResult.files,
        extractionResult.tempFolder,
      );

      console.log("Configuration data extracted successfully");
      console.log(
        `- Presets: ${configData.presets ? configData.presets.length : 0}`,
      );
      console.log(
        `- Fields: ${configData.fields ? configData.fields.length : 0}`,
      );
      console.log(`- Icons: ${configData.icons ? configData.icons.length : 0}`);
    } catch (error) {
      console.error("Error extracting configuration data:", error);
      return {
        success: false,
        message: `Failed to extract configuration data: ${error instanceof Error ? error.message : String(error)}`,
      };
    }

    // Apply configuration to spreadsheet
    console.log("Applying configuration to spreadsheet...");
    try {
      applyConfigurationToSpreadsheet(configData);
      console.log("Configuration applied to spreadsheet successfully");
    } catch (error) {
      console.error("Error applying configuration to spreadsheet:", error);
      return {
        success: false,
        message: `Failed to apply configuration to spreadsheet: ${error instanceof Error ? error.message : String(error)}`,
      };
    }

    // Clean up
    console.log("Cleaning up...");
    try {
      extractionResult.tempFolder.setTrashed(true);
      console.log("Temporary folder deleted");
    } catch (error) {
      console.warn("Error cleaning up temporary folder:", error);
    }

    // Return success
    console.log("Import completed successfully");
    const fileType = fileName.endsWith(".comapeocat") ? "CoMapeo" : "Mapeo";
    return {
      success: true,
      message: `${fileType} category file imported successfully`,
      details: {
        presets: configData.presets ? configData.presets.length : 0,
        fields: configData.fields ? configData.fields.length : 0,
        icons: configData.icons ? configData.icons.length : 0,
        languages: Object.keys(configData.messages).length,
      },
      warnings: extractionResult.validationWarnings || [],
    };
  } catch (error) {
    console.error("Error importing configuration file:", error);
    return {
      success: false,
      message: error instanceof Error ? error.message : String(error),
    };
  }
}
