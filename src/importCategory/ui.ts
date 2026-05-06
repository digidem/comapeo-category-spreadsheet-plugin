/**
 * UI functions for the import category functionality.
 * This file contains functions related to the user interface.
 */

/**
 * Creates the HTML for the import category dialog.
 * @returns HTML string for the dialog
 */
function createImportCategoryHtml(): string {
  const title = importCategoryDialogTexts[activeLocale].title;
  const messages = importCategoryDialogTexts[activeLocale].message
    .map((msg) => "<p>" + msg + "</p>")
    .join("");
  const buttonText = importCategoryDialogTexts[activeLocale].buttonText;

  return (
    "<!DOCTYPE html>" +
    "<html>" +
    "<head>" +
    '  <base target="_top">' +
    "  <style>" +
    "    body { font-family: 'Roboto', sans-serif; line-height: 1.6; color: #e0e0e0; background: linear-gradient(135deg, #1a1a1a, #2c2c2c); padding: 20px; }" +
    "    h1 { color: #6d44d9; font-size: 1.2em; margin-bottom: 20px; text-align: center; }" +
    "    p { margin-bottom: 15px; text-align: center; }" +
    "    .file-upload-container { margin: 20px 0; text-align: center; }" +
    "    .file-upload-label { display: inline-block; padding: 12px 20px; background: linear-gradient(45deg, #330B9E, #6d44d9); color: white; border-radius: 50px; cursor: pointer; transition: all 0.3s ease; }" +
    "    .file-upload-label:hover { background: linear-gradient(45deg, #4A0ED6, #8a67e8); transform: translateY(-2px); }" +
    "    .upload-icon { margin-right: 8px; font-size: 1.2em; }" +
    "    .file-info { margin: 15px 0; text-align: center; }" +
    "    .upload-status { margin: 15px 0; text-align: center; }" +
    "    .success { color: #4CAF50; font-weight: bold; }" +
    "    .error { color: #F44336; font-weight: bold; }" +
    "    .progress-container { margin: 20px auto; max-width: 400px; opacity: 0; animation: fadeIn 0.3s forwards; }" +
    "    .progress-header { display: flex; justify-content: space-between; margin-bottom: 8px; color: #b0b0b0; font-size: 0.9em; }" +
    "    .progress-bar-container { width: 100%; height: 24px; background: rgba(255,255,255,0.1); border-radius: 12px; overflow: hidden; box-shadow: inset 0 2px 4px rgba(0,0,0,0.3); }" +
    "    .progress-bar { height: 100%; background: linear-gradient(90deg, #330B9E, #6d44d9, #8a67e8); border-radius: 12px; transition: width 0.3s ease; box-shadow: 0 0 10px rgba(109, 68, 217, 0.5); width: 0%; }" +
    "    .progress-detail { margin-top: 8px; text-align: center; color: #888; font-size: 0.85em; min-height: 20px; }" +
    "    @keyframes fadeIn { to { opacity: 1; } }" +
    "  </style>" +
    "</head>" +
    "<body>" +
    "  <h1>" +
    title +
    "</h1>" +
    "  " +
    messages +
    "" +
    "  " +
    '  <div class="file-upload-container">' +
    '    <label for="file" class="file-upload-label">' +
    '      <span class="upload-icon">📂</span>' +
    '      <span class="upload-text">' +
    buttonText +
    "</span>" +
    "    </label>" +
    '    <input type="file" id="file" name="file" accept=".comapeocat,.zip,.mapeosettings" style="display: none;" onchange="handleFileSelect()">' +
    "  </div>" +
    '  <div id="file-info" class="file-info"></div>' +
    '  <div id="progress-container" class="progress-container" style="display: none;">' +
    '    <div class="progress-header">' +
    '      <span id="progress-stage">Processing...</span>' +
    '      <span id="progress-percent">0%</span>' +
    "    </div>" +
    '    <div class="progress-bar-container">' +
    '      <div id="progress-bar" class="progress-bar"></div>' +
    "    </div>" +
    '    <div id="progress-detail" class="progress-detail"></div>' +
    "  </div>" +
    '  <div id="upload-status" class="upload-status"></div>' +
    "  " +
    "  <script>" +
    "    function handleFileSelect() {" +
    '      const fileInput = document.getElementById("file");' +
    '      const fileInfo = document.getElementById("file-info");' +
    '      const uploadStatus = document.getElementById("upload-status");' +
    '      const progressContainer = document.getElementById("progress-container");' +
    "      " +
    "      if (fileInput.files.length > 0) {" +
    "        const file = fileInput.files[0];" +
    '        fileInfo.innerHTML = "<p>Selected file: " + file.name + " (" + (file.size / 1024).toFixed(2) + " KB)</p>";' +
    '        progressContainer.style.display = "block";' +
    '        uploadStatus.innerHTML = "";' +
    "        " +
    "        // Show indeterminate progress" +
    '        document.getElementById("progress-bar").style.width = "50%";' +
    '        document.getElementById("progress-stage").textContent = "Processing...";' +
    '        document.getElementById("progress-percent").textContent = "";' +
    '        document.getElementById("progress-detail").textContent = "Importing configuration";' +
    "        " +
    "        // Read the file and convert to base64" +
    "        const reader = new FileReader();" +
    "        reader.onload = function(e) {" +
    '          const base64data = e.target.result.split(",")[1];' +
    "          // Note: Progress callbacks cannot work across Apps Script client/server boundary" +
    "          google.script.run" +
    "            .withSuccessHandler(onSuccess)" +
    "            .withFailureHandler(onFailure)" +
    "            .processImportedCategoryFile(file.name, base64data);" +
    "        };" +
    "        reader.readAsDataURL(file);" +
    "      }" +
    "    }" +
    "    " +
    "    function onSuccess(result) {" +
    '      const uploadStatus = document.getElementById("upload-status");' +
    '      const progressContainer = document.getElementById("progress-container");' +
    "      " +
    "      // Hide progress, show success" +
    '      progressContainer.style.display = "none";' +
    "      " +
    "      // Get success message from dialog texts" +
    "      const successMsg = " + JSON.stringify(importCategoryDialogTexts) + "['" + activeLocale + "'].successMessage;" +
    "      " +
    "      // Build detailed success message" +
    "      let message = \"<div style='text-align: left; max-width: 600px; margin: 0 auto;'>\";" +
    "      message += \"<p class='success' style='text-align: center; font-size: 1.2em; margin-bottom: 20px;'>✓ \" + successMsg.title + \"</p>\";" +
    "      message += \"<div style='background: rgba(255,255,255,0.05); padding: 15px; border-radius: 8px; font-size: 0.9em; line-height: 1.6;'>\";" +
    "      successMsg.dropdownInfo.forEach(function(line) {" +
    "        if (line === '') {" +
    "          message += \"<br>\";" +
    "        } else {" +
    "          message += \"<p style='margin: 5px 0;'>\" + line + \"</p>\";" +
    "        }" +
    "      });" +
    "      message += \"</div>\";" +
    "      message += \"<div style='text-align: center; margin-top: 20px;'>\";" +
    "      message += \"<button onclick='google.script.host.close()' style='padding: 10px 24px; background: linear-gradient(45deg, #330B9E, #6d44d9); color: white; border: none; border-radius: 50px; cursor: pointer; font-size: 1em; transition: all 0.3s ease;' onmouseover='this.style.background=\\\"linear-gradient(45deg, #4A0ED6, #8a67e8)\\\"' onmouseout='this.style.background=\\\"linear-gradient(45deg, #330B9E, #6d44d9)\\\">Close</button>\";" +
    "      message += \"<p style='color: #888; font-size: 0.85em; margin-top: 10px;'>Auto-closing in 90 seconds...</p>\";" +
    "      message += \"</div>\";" +
    "      message += \"</div>\";" +
    "      " +
    "      uploadStatus.innerHTML = message;" +
    "      " +
    "      setTimeout(function() {" +
    "        google.script.host.close();" +
    "      }, 90000);" +
    "    }" +
    "    " +
    "    function onFailure(error) {" +
    '      const uploadStatus = document.getElementById("upload-status");' +
    '      const progressContainer = document.getElementById("progress-container");' +
    "      " +
    "      // Hide progress, show error" +
    '      progressContainer.style.display = "none";' +
    '      uploadStatus.innerHTML = "<p class=\'error\'>✗ Error: " + error.message + "</p>";' +
    "    }" +
    "  </script>" +
    "</body>" +
    "</html>"
  );
}
