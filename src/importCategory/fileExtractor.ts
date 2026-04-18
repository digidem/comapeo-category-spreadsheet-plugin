/**
 * Utilities for extracting and validating uploaded configuration files
 */

/**
 * Required files and directories that should be present in a valid configuration
 */
const REQUIRED_FILES = [
  "metadata.json", // or package.json
  "presets.json", // or presets/ directory
  "translations.json", // can be created if missing
  "icons.svg",
];

/**
 * Optional files that may be present in a valid configuration
 */
const OPTIONAL_FILES = [
  "icons", // directory
  "icons.json",
  "icons.png",
  "style.css",
  "VERSION",
  "defaults.json",
];

/**
 * Result of file extraction and validation
 */
interface ExtractionResult {
  success: boolean;
  message: string;
  files?: GoogleAppsScript.Base.Blob[];
  validationErrors?: string[];
  validationWarnings?: string[];
  tempFolder?: GoogleAppsScript.Drive.Folder;
}

/**
 * Options for extraction process
 */
interface ExtractionOptions {
  onProgress?: (stage: string, percent: number) => void;
}

function extractJsonArchive(
  fileBlob: GoogleAppsScript.Base.Blob,
  tempFolder: GoogleAppsScript.Drive.Folder,
): GoogleAppsScript.Base.Blob[] {
  const extractedFiles: GoogleAppsScript.Base.Blob[] = [];
  const raw = fileBlob.getDataAsString();
  const jsonData = JSON.parse(raw);

  const createAndStore = (
    content: any,
    name: string,
    mimeType = "application/json",
  ) => {
    if (!content) {
      return;
    }
    const blob = Utilities.newBlob(
      typeof content === "string" ? content : JSON.stringify(content, null, 2),
      mimeType,
      name,
    );
    tempFolder.createFile(blob);
    extractedFiles.push(blob);
  };

  if (jsonData.metadata) {
    createAndStore(jsonData.metadata, "metadata.json");
  }

  if (jsonData.presets || jsonData.fields) {
    const payload: Record<string, any> = {
      presets: jsonData.presets || {},
    };
    if (jsonData.fields) {
      payload.fields = jsonData.fields;
    }
    createAndStore(payload, "presets.json");
  }

  if (jsonData.translations) {
    createAndStore(jsonData.translations, "translations.json");
  }

  if (jsonData.icons) {
    try {
      let iconsSvgContent = "";
      if (typeof jsonData.icons === "string") {
        iconsSvgContent = jsonData.icons;
      } else if (jsonData.icons.svg) {
        iconsSvgContent = jsonData.icons.svg;
      } else if (Array.isArray(jsonData.icons)) {
        iconsSvgContent = createIconSpriteFromArray(jsonData.icons);
      }

      if (iconsSvgContent) {
        createAndStore(iconsSvgContent, "icons.svg", "image/svg+xml");
      }
    } catch (iconError) {
      console.warn("Failed to generate icons.svg from JSON archive", iconError);
    }
  }

  return extractedFiles;
}

/**
 * Escapes a string for safe use as an XML attribute value.
 * Covers &, <, >, and " to prevent injection via icon names
 * containing XML metacharacters (e.g. "fish & shell").
 */
function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * Creates an SVG sprite string from an array of icon definitions
 */
function createIconSpriteFromArray(icons: Array<{ name?: string; svg?: string }>): string {
  const safeIcons = Array.isArray(icons) ? icons : [];
  const extractSymbolMarkup = (icon: { name?: string; svg?: string }): string | null => {
    if (!icon || !icon.name || !icon.svg) {
      return null;
    }

    const trimmed = icon.svg.trim();
    if (!trimmed) {
      return null;
    }

    const hasWrappedRoot = /^<\s*(svg|symbol)\b/i.test(trimmed);
    const wrapped = hasWrappedRoot ? trimmed : `<svg>${trimmed}</svg>`;

    try {
      const doc = XmlService.parse(wrapped);
      const rootEl = doc.getRootElement();
      const rootName = rootEl.getName().toLowerCase();
      const viewBox = rootEl.getAttribute("viewBox")?.getValue();
      const body = rootName === "symbol" || rootName === "svg"
        ? trimmed.replace(/^<\s*(svg|symbol)\b[\s\S]*?>/i, "").replace(/<\s*\/\s*(svg|symbol)\s*>$/i, "").trim()
        : trimmed;
      // Preserve namespace attributes from the original root element (e.g.
      // xmlns:xlink) so that inner markup referencing prefixed attributes
      // remains valid after the root element is replaced with <symbol>.
      const nsAttrs: string[] = [];
      const rootAttrs = rootEl.getAttributes();
      for (const attr of rootAttrs) {
        const attrName = attr.getName();
        const attrPrefix = attr.getNamespace()?.getPrefix();
        // Keep xmlns:xxx declarations (namespace prefix declarations)
        if (attrPrefix === "xmlns" || attrName === "xmlns" || attrName.startsWith("xmlns:")) {
          const ns = attr.getNamespace();
          const fullAttrName = ns?.getPrefix() ? `${ns.getPrefix()}:${attrName}` : attrName;
          nsAttrs.push(`${fullAttrName}="${attr.getValue().replace(/"/g, "&quot;")}"`);
        }
      }
      const attrs = viewBox ? ` viewBox="${viewBox.replace(/"/g, "&quot;")}"` : "";
      const nsAttrStr = nsAttrs.length > 0 ? ` ${nsAttrs.join(" ")}` : "";
      return `<symbol id="${escapeXmlAttr(icon.name)}"${attrs}${nsAttrStr}>${body}</symbol>`;
    } catch (error) {
      console.warn("Failed to parse icon SVG with XmlService, falling back to raw string:", error);
      // Keep the fallback XML-safe: escape the raw body as text content so one
      // malformed icon cannot break XmlService.parse() on the entire sprite.
      const body = trimmed
        .replace(/^<\s*svg\b[\s\S]*?>/i, "")
        .replace(/<\s*\/\s*svg\s*>$/i, "")
        .trim();
      const escapedBody = body
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
      return `<symbol id="${escapeXmlAttr(icon.name)}">${escapedBody}</symbol>`;
    }
  };

  const symbols = safeIcons
    .map(extractSymbolMarkup)
    .filter((symbol): symbol is string => Boolean(symbol));

  return `<svg xmlns="http://www.w3.org/2000/svg">${symbols.join("")}</svg>`;
}

/**
 * Maximum file size for imports (100MB)
 */
const MAX_FILE_SIZE = 100 * 1024 * 1024;

/**
 * Validates a file path to prevent path traversal attacks
 *
 * @param path - The file path to validate
 * @returns true if valid, throws error if invalid
 * @throws Error if path contains path traversal sequences
 */
function validateFilePathSecurity(path: string): boolean {
  // Check for path traversal attempts
  if (path.includes("../") || path.includes("..\\")) {
    throw new Error(`Invalid file path detected (path traversal attempt): ${path}`);
  }

  // Check for absolute paths (Unix and Windows)
  if (path.startsWith("/") || /^[A-Za-z]:/.test(path)) {
    throw new Error(`Invalid file path detected (absolute path): ${path}`);
  }

  return true;
}

/**
 * Extracts and validates the content of an uploaded file
 * @param fileName - Name of the uploaded file
 * @param fileBlob - The file blob to extract
 * @param options - Optional extraction options
 * @returns Result of extraction and validation
 */
function extractAndValidateFile(
  fileName: string,
  fileBlob: GoogleAppsScript.Base.Blob,
  options?: ExtractionOptions,
): ExtractionResult {
  // Helper function to report progress
  const reportProgress = (stage: string, percent: number) => {
    if (options?.onProgress) {
      options.onProgress(stage, percent);
    }
  };

  try {
    // Validate file size to prevent memory exhaustion
    const fileSize = fileBlob.getBytes().length;
    const fileSizeMB = (fileSize / 1024 / 1024).toFixed(2);
    console.log(`File size: ${fileSizeMB} MB`);

    if (fileSize > MAX_FILE_SIZE) {
      const maxSizeMB = (MAX_FILE_SIZE / 1024 / 1024).toFixed(0);
      const errorMessage = `File too large: ${fileSizeMB} MB (maximum: ${maxSizeMB} MB)`;
      console.error(errorMessage);
      return {
        success: false,
        message: errorMessage,
        validationErrors: [errorMessage],
      };
    }

    reportProgress("Creating temporary folder", 5);
    // Create a temporary folder for extraction
    const tempFolderName = "Config_Import_Temp_" + new Date().getTime();
    const tempFolder = DriveApp.createFolder(tempFolderName);

    // Determine file type and extract accordingly
    const fileExtension = fileName.split(".").pop()?.toLowerCase();
    reportProgress("Determining file type", 10);

    let extractedFiles: GoogleAppsScript.Base.Blob[] = [];

    if (fileExtension === "comapeocat" || fileExtension === "zip") {
      reportProgress("Extracting ZIP file", 15);
      // Extract zip file
      try {
        extractedFiles = extractZipFile(fileBlob, tempFolder, (percent) => {
          reportProgress("Extracting ZIP file", 15 + Math.round(percent * 0.5)); // Map 0-100 to 15-65
        });
      } catch (zipError) {
        console.warn("ZIP extraction failed, attempting JSON fallback", zipError);
        try {
          extractedFiles = extractJsonArchive(fileBlob, tempFolder);
          reportProgress("Processed JSON archive", 40);
        } catch (jsonError) {
          console.error("JSON fallback also failed", jsonError);
          tempFolder.setTrashed(true);
          return {
            success: false,
            message:
              "Failed to extract configuration: " +
              (jsonError instanceof Error ? jsonError.message : String(jsonError)),
          };
        }
      }
    } else if (fileExtension === "mapeosettings" || fileExtension === "tar") {
      reportProgress("Extracting TAR file", 15);
      // Extract tar file
      extractedFiles = extractTarArchiveInternal(fileBlob, tempFolder, (percent) => {
        reportProgress("Extracting TAR file", 15 + Math.round(percent * 0.5)); // Map 0-100 to 15-65
      });
    } else {
      // Clean up and return error
      tempFolder.setTrashed(true);
      return {
        success: false,
        message:
          "Unsupported file format. Please upload a .comapeocat, .zip, .mapeosettings, or .tar file.",
      };
    }

    // Validate all extracted file paths for security
    reportProgress("Validating file paths", 70);
    const pathErrors: string[] = [];
    for (const file of extractedFiles) {
      const extractedFileName = file.getName();
      try {
        validateFilePathSecurity(extractedFileName);
      } catch (pathError) {
        console.error(`Path validation failed for ${extractedFileName}: ${pathError.message}`);
        pathErrors.push(pathError.message);
      }
    }

    if (pathErrors.length > 0) {
      // Clean up temp folder
      tempFolder.setTrashed(true);
      return {
        success: false,
        message: `Security validation failed: ${pathErrors.length} invalid file path(s) detected`,
        validationErrors: pathErrors,
      };
    }

    // Validate extracted structure before returning
    reportProgress("Validating extracted structure", 80);
    const validation = validateExtractedFiles(extractedFiles, (percent) => {
      reportProgress("Validating extracted structure", 80 + Math.round(percent * 0.1));
    });

    if (!validation.success) {
      const errorMessage =
        validation.errors && validation.errors.length > 0
          ? validation.errors.join("; ")
          : "Unknown validation error";
      console.error("Validation failed:", errorMessage);
      try {
        tempFolder.setTrashed(true);
      } catch (cleanupError) {
        console.warn("Failed to clean up temp folder after validation failure", cleanupError);
      }
      return {
        success: false,
        message: `Validation failed: ${errorMessage}`,
        validationErrors: validation.errors,
      };
    }

    return {
      success: true,
      message: "File extracted successfully",
      files: extractedFiles,
      tempFolder: tempFolder,
      validationWarnings: validation.warnings,
    };
  } catch (error) {
    console.error("Error extracting file:", error);
    return {
      success: false,
      message:
        "Error extracting file: " +
        (error instanceof Error ? error.message : String(error)),
    };
  }
}

/**
 * Extracts a zip file
 * @param fileBlob - The zip file blob
 * @param tempFolder - Folder to save extracted files
 * @param progressCallback - Optional callback for progress updates
 * @returns Array of extracted file blobs
 */
function extractZipFile(
  fileBlob: GoogleAppsScript.Base.Blob,
  tempFolder: GoogleAppsScript.Drive.Folder,
  progressCallback?: (percent: number) => void,
): GoogleAppsScript.Base.Blob[] {
  try {
    // Helper function to report progress
    const reportProgress = (percent: number) => {
      if (progressCallback) {
        progressCallback(percent);
      }
    };

    reportProgress(5);
    // Ensure we have the correct content type for zip files
    const zipBlob = fileBlob.setContentType("application/zip");

    // Log file details for debugging
    console.log("Extracting zip file:", {
      name: zipBlob.getName(),
      size: zipBlob.getBytes().length,
      contentType: zipBlob.getContentType(),
    });
    reportProgress(10);

    // Use built-in unzip utility
    let unzippedFiles: GoogleAppsScript.Base.Blob[];
    reportProgress(15);
    console.log("Starting unzip operation...");
    unzippedFiles = Utilities.unzip(zipBlob);
    reportProgress(40);
    console.log("Unzip operation completed successfully");

    // Log the extracted files
    console.log(
      "Extracted files:",
      unzippedFiles.map((f) => f.getName()),
    );
    reportProgress(50);

    // Save files to temp folder for reference
    const savedFiles: GoogleAppsScript.Base.Blob[] = [];
    const totalFiles = unzippedFiles.length;

    console.log(`Processing ${totalFiles} extracted files...`);
    unzippedFiles.forEach((file, index) => {
      try {
        const fileName = file.getName();
        console.log(`Processing file ${index + 1}/${totalFiles}: ${fileName}`);

        // Update progress based on file processing (50-90%)
        const fileProgress = 50 + Math.round((index / totalFiles) * 40);
        reportProgress(fileProgress);

        // Skip directories in the zip
        if (!fileName.endsWith("/")) {
          // Create parent directories if needed
          if (fileName.includes("/")) {
            const dirPath = fileName.substring(0, fileName.lastIndexOf("/"));
            createNestedFolders(tempFolder, dirPath);
          }

          // Save the file
          const savedFile = tempFolder.createFile(file);
          console.log("Saved file:", savedFile.getName());
          savedFiles.push(file);
        }
        // else {
        //   // Create directories
        //   const dirPath = fileName.slice(0, -1); // Remove trailing slash
        //   createNestedFolders(tempFolder, dirPath);
        //   savedFiles.push(file);
        // }
      } catch (fileError) {
        console.warn("Error saving extracted file:", file.getName(), fileError);
        // Continue with other files even if one fails
      }
    });

    reportProgress(95);
    console.log("All files processed successfully");
    reportProgress(100);
    return unzippedFiles;
  } catch (error) {
    console.error("Error extracting zip file:", error);
    throw new Error(
      "Failed to extract zip file: " +
        (error instanceof Error ? error.message : String(error)),
    );
  }
}

/**
 * Extracts a tar file
 * @param fileBlob - The tar file blob
 * @param tempFolder - Folder to save extracted files
 * @param progressCallback - Optional callback for progress updates
 * @param debugMode - Optional debug mode for additional logging
 * @returns Array of extracted file blobs
 */
function extractTarArchiveInternal(
  fileBlob: GoogleAppsScript.Base.Blob,
  tempFolder: GoogleAppsScript.Drive.Folder,
  progressCallback?: (percent: number) => void,
  debugMode?: boolean,
): GoogleAppsScript.Base.Blob[] {
  // Log file details at the start
  console.log("TAR extraction started:", {
    fileName: fileBlob.getName(),
    fileSize: fileBlob.getBytes().length,
    contentType: fileBlob.getContentType(),
    tempFolderName: tempFolder.getName(),
    debugMode: !!debugMode,
  });

  try {
    // Helper function to report progress
    const reportProgress = (percent: number) => {
      if (progressCallback) {
        progressCallback(percent);
      }
    };

    reportProgress(5);
    console.log("Starting TAR file extraction...");

    // Try to parse as JSON first (some .mapeosettings files are just JSON)
    try {
      reportProgress(10);
      console.log("Attempting to parse as JSON...");
      const fileContent = fileBlob.getDataAsString();
      reportProgress(20);
      const jsonData = JSON.parse(fileContent);
      reportProgress(30);
      console.log("Successfully parsed as JSON");

      // If we got here, it's valid JSON
      const extractedFiles: GoogleAppsScript.Base.Blob[] = [];

      // Create metadata.json
      if (jsonData.metadata) {
        console.log("Creating metadata.json...");
        const metadataBlob = Utilities.newBlob(
          JSON.stringify(jsonData.metadata, null, 2),
          "application/json",
          "metadata.json",
        );
        tempFolder.createFile(metadataBlob);
        extractedFiles.push(metadataBlob);
      }

      // Create presets.json
      if (jsonData.presets) {
        console.log("Creating presets.json...");
        const presetsBlob = Utilities.newBlob(
          JSON.stringify({ presets: jsonData.presets }, null, 2),
          "application/json",
          "presets.json",
        );
        tempFolder.createFile(presetsBlob);
        extractedFiles.push(presetsBlob);
      }

      // Create translations.json
      if (jsonData.translations) {
        console.log("Creating translations.json...");
        const translationsBlob = Utilities.newBlob(
          JSON.stringify(jsonData.translations, null, 2),
          "application/json",
          "translations.json",
        );
        tempFolder.createFile(translationsBlob);
        extractedFiles.push(translationsBlob);
      }

      // Handle embedded icons structures when present
      if (jsonData.icons) {
        console.log("Processing embedded icons data...");
        try {
          let iconsSvgContent = "";

          if (typeof jsonData.icons === "string") {
            iconsSvgContent = jsonData.icons;
          } else if (jsonData.icons.svg) {
            iconsSvgContent = jsonData.icons.svg;
          } else if (Array.isArray(jsonData.icons)) {
            console.log(
              `Found ${jsonData.icons.length} icon objects, building sprite...`,
            );
            iconsSvgContent = createIconSpriteFromArray(jsonData.icons);
          }

          if (iconsSvgContent) {
            const iconsSvgBlob = Utilities.newBlob(
              iconsSvgContent,
              "image/svg+xml",
              "icons.svg",
            );
            tempFolder.createFile(iconsSvgBlob);
            extractedFiles.push(iconsSvgBlob);
            console.log("Embedded icons converted to icons.svg");
          } else {
            console.warn("Icons data present but no SVG content could be derived");
          }
        } catch (iconError) {
          console.error("Failed to process embedded icons:", iconError);
        }
      }

      reportProgress(90);
      console.log("JSON extraction completed successfully");
      reportProgress(100);
      return extractedFiles;
    } catch (jsonError) {
      // Not a JSON file, must be an actual tar file
      console.warn("Not a JSON file, attempting to extract as tar:", jsonError);
      reportProgress(30);

      // Implement basic tar extraction
      console.log("Implementing tar extraction...");

      // TAR file format constants
      const HEADER_SIZE = 512;
      const BLOCK_SIZE = 512;
      const NAME_OFFSET = 0;
      const NAME_LENGTH = 100;
      const SIZE_OFFSET = 124;
      const SIZE_LENGTH = 12;
      const TYPE_FLAG_OFFSET = 156;
      const MAGIC_OFFSET = 257;
      const MAGIC_LENGTH = 6;

      // Get the raw bytes of the tar file
      const bytes = fileBlob.getBytes();
      console.log(`TAR file size: ${bytes.length} bytes`);
      reportProgress(35);

      // Array to store extracted files
      const extractedFiles: GoogleAppsScript.Base.Blob[] = [];

      // Process the tar file
      let position = 0;
      let fileCount = 0;
      const processedFiles = new Set<string>();

      while (position + HEADER_SIZE <= bytes.length) {
        // Read header block
        const headerBlock = bytes.slice(position, position + HEADER_SIZE);

        // Check for end of archive (empty block)
        let isEmpty = true;
        for (let i = 0; i < HEADER_SIZE; i++) {
          if (headerBlock[i] !== 0) {
            isEmpty = false;
            break;
          }
        }

        if (isEmpty) {
          console.log("Reached end of archive (empty block)");
          break;
        }

        // Check magic number "ustar"
        const magicBytes = headerBlock.slice(
          MAGIC_OFFSET,
          MAGIC_OFFSET + MAGIC_LENGTH,
        );
        const magic = String.fromCharCode.apply(null, magicBytes);

        if (magic.indexOf("ustar") !== 0) {
          console.warn(
            `Invalid magic number at position ${position}, skipping block`,
          );
          position += BLOCK_SIZE;
          continue;
        }

        // Extract file name (null-terminated C string)
        const nameBytes = headerBlock.slice(
          NAME_OFFSET,
          NAME_OFFSET + NAME_LENGTH,
        );
        let fileName = "";
        for (let i = 0; i < nameBytes.length; i++) {
          if (nameBytes[i] === 0) break;
          fileName += String.fromCharCode(nameBytes[i]);
        }

        // Skip if we've already processed this file (avoid duplicates)
        if (processedFiles.has(fileName)) {
          console.log(`Skipping duplicate file: ${fileName}`);
          position += BLOCK_SIZE;
          continue;
        }

        // Extract file size (octal string)
        const sizeBytes = headerBlock.slice(
          SIZE_OFFSET,
          SIZE_OFFSET + SIZE_LENGTH,
        );
        let sizeStr = "";
        for (let i = 0; i < sizeBytes.length; i++) {
          if (sizeBytes[i] === 0) break;
          sizeStr += String.fromCharCode(sizeBytes[i]);
        }

        // Parse size as octal
        const fileSize = parseInt(sizeStr.trim(), 8);

        // Get type flag (0 or '0' = regular file, 5 = directory)
        const typeFlag = headerBlock[TYPE_FLAG_OFFSET];
        const isDirectory = typeFlag === 53; // ASCII '5'

        console.log(
          `Found ${isDirectory ? "directory" : "file"}: ${fileName}, size: ${fileSize}`,
        );

        // Process core config files and anything in icons/ directory
        const isCoreFile =
          fileName === "metadata.json" ||
          fileName === "presets.json" ||
          fileName === "translations.json" ||
          fileName === "icons.svg";

        // Check if file is in icons/ directory (handles both "icons/file.png" and "path/icons/file.png")
        const isIconFile =
          fileName.includes("/icons/") || fileName.startsWith("icons/");

        const shouldExtract = isCoreFile || isIconFile;

        if (!isDirectory && shouldExtract) {
          // Extract file data
          const fileData = bytes.slice(
            position + HEADER_SIZE,
            position + HEADER_SIZE + fileSize,
          );

          // Determine MIME type based on file extension
          let mimeType = "application/json";
          if (fileName.endsWith(".svg")) {
            mimeType = "image/svg+xml";
          } else if (fileName.endsWith(".png")) {
            mimeType = "image/png";
          }

          // Create file blob with full path preserved
          const fileBlob = Utilities.newBlob(fileData, mimeType, fileName);

          // For icon files, create nested folder structure
          if (isIconFile && fileName.includes("/")) {
            // Extract the directory path
            const lastSlashIndex = fileName.lastIndexOf("/");
            const dirPath = fileName.substring(0, lastSlashIndex);

            console.log(`Creating nested folder structure: ${dirPath}`);

            // Create nested folders if they don't exist
            let currentFolder = tempFolder;
            const parts = dirPath.split("/");

            for (const part of parts) {
              if (part) {
                const subfolders = currentFolder.getFoldersByName(part);
                if (subfolders.hasNext()) {
                  currentFolder = subfolders.next();
                } else {
                  currentFolder = currentFolder.createFolder(part);
                }
              }
            }

            // Save file to the nested folder with just the filename
            const justFileName = fileName.substring(lastSlashIndex + 1);
            console.log(
              `Saving icon file: ${justFileName} to folder: ${currentFolder.getName()}`,
            );
            currentFolder.createFile(
              Utilities.newBlob(fileData, mimeType, justFileName),
            );
          } else {
            // Save core files directly to temp folder
            tempFolder.createFile(fileBlob);
          }

          extractedFiles.push(fileBlob);
        }

        // Calculate how many blocks to skip
        const dataBlocks = Math.ceil(fileSize / BLOCK_SIZE);
        const dataSize = dataBlocks * BLOCK_SIZE;
        position += HEADER_SIZE + dataSize;

        // Mark file as processed
        processedFiles.add(fileName);
        fileCount++;

        // Update progress
        reportProgress(35 + Math.min(55, Math.round((fileCount / 50) * 55)));
      }

      console.log(`Processed ${fileCount} files from tar archive`);
      reportProgress(100);
      console.log("TAR extraction completed successfully");
      return extractedFiles;
    }
  } catch (error) {
    console.error("Error extracting tar file:", error);
    throw new Error(
      "Failed to extract tar file: " +
        (error instanceof Error ? error.message : String(error)),
    );
  }
}

/**
 * Creates nested folders in Drive
 * @param parentFolder - The parent folder
 * @param path - Path to create
 * @returns The created folder
 */
function createNestedFolders(
  parentFolder: GoogleAppsScript.Drive.Folder,
  path: string,
): GoogleAppsScript.Drive.Folder {
  const parts = path.split("/");
  let currentFolder = parentFolder;

  for (const part of parts) {
    if (part) {
      // Check if folder already exists
      const folderIterator = currentFolder.getFoldersByName(part);
      if (folderIterator.hasNext()) {
        currentFolder = folderIterator.next();
      } else {
        currentFolder = currentFolder.createFolder(part);
      }
    }
  }

  return currentFolder;
}

/**
 * Backwards-compatible wrapper for tests and legacy code paths.
 */
function extractTarFile(
  fileBlob: GoogleAppsScript.Base.Blob,
  tempFolder: GoogleAppsScript.Drive.Folder,
  progressCallback?: (percent: number) => void,
  debugMode?: boolean,
): GoogleAppsScript.Base.Blob[] {
  return extractTarArchiveInternal(fileBlob, tempFolder, progressCallback, debugMode);
}

/**
 * Validates the extracted files against required structure
 * @param files - Array of extracted file blobs
 * @param progressCallback - Optional callback for progress updates
 * @returns Validation result
 */
function validateExtractedFiles(
  files: GoogleAppsScript.Base.Blob[],
  progressCallback?: (percent: number) => void,
): { success: boolean; errors?: string[]; warnings?: string[] } {
  // Helper function to report progress
  const reportProgress = (percent: number) => {
    if (progressCallback) {
      progressCallback(percent);
    }
  };

  const errors: string[] = [];
  const warnings: string[] = [];
  const foundFiles = new Set<string>();
  const foundDirs = new Set<string>();

  console.log("Validating extracted files:", files.length);
  reportProgress(10);

  // Check for each file in the extracted content
  files.forEach((file, index) => {
    const fileName = file.getName();
    console.log(`Checking file ${index + 1}/${files.length}: ${fileName}`);

    // Update progress (10-40%)
    const fileProgress = 10 + Math.round((index / files.length) * 30);
    reportProgress(fileProgress);

    // Handle directories (they end with '/' in zip files)
    if (fileName.endsWith("/")) {
      const dirName = fileName.slice(0, -1);
      foundDirs.add(dirName);

      // Also add the top-level directory to foundFiles
      const topLevelDir = dirName.split("/")[0];
      foundFiles.add(topLevelDir);
      console.log("Found directory:", dirName, "(top-level:", topLevelDir, ")");
    } else {
      // Handle regular files
      foundFiles.add(fileName);

      if (fileName.includes("/")) {
        // For nested files, add the directory structure
        const parts = fileName.split("/");
        let currentPath = "";

        for (let i = 0; i < parts.length - 1; i++) {
          if (currentPath) {
            currentPath += "/" + parts[i];
          } else {
            currentPath = parts[i];
          }

          foundDirs.add(currentPath);

          // Add top-level directory to foundFiles
          if (i === 0) {
            foundFiles.add(parts[0]);
          }
        }

        console.log(
          "Found nested file:",
          fileName,
          "in directories:",
          Array.from(foundDirs),
        );
      } else {
        // For top-level files
        console.log("Found top-level file:", fileName);
      }
    }
  });

  console.log("Found files:", Array.from(foundFiles));
  console.log("Found directories:", Array.from(foundDirs));
  reportProgress(50);

  // Check if we have a presets directory
  const hasPresetsDir = foundDirs.has("presets");
  console.log("Has presets directory:", hasPresetsDir);

  // Check for missing required files
  console.log("Checking for required files...");
  REQUIRED_FILES.forEach((requiredFile, index) => {
    // Update progress (50-60%)
    const requiredProgress =
      50 + Math.round((index / REQUIRED_FILES.length) * 10);
    reportProgress(requiredProgress);

    if (!foundFiles.has(requiredFile)) {
      // Special case for metadata.json - some files might have meta.json instead
      if (requiredFile === "metadata.json" && foundFiles.has("meta.json")) {
        console.log("Found meta.json instead of metadata.json - acceptable");
      }
      // Special case for presets.json - we might have individual preset files
      else if (requiredFile === "presets.json" && hasPresetsDir) {
        console.log(
          "Found presets directory instead of presets.json - acceptable",
        );
      }
      // Special case for translations.json - we might not have translations
      else if (requiredFile === "translations.json") {
        console.log("Missing translations.json - will create empty one");
        warnings.push("Missing translations.json file - created empty one");
      } else {
        errors.push(`Missing required file or directory: ${requiredFile}`);
        console.warn(`Missing required file: ${requiredFile}`);
      }
    }
  });

  // Check for missing optional files
  console.log("Checking for optional files...");
  reportProgress(65);
  OPTIONAL_FILES.forEach((optionalFile, index) => {
    // Update progress (65-75%)
    const optionalProgress =
      65 + Math.round((index / OPTIONAL_FILES.length) * 10);
    reportProgress(optionalProgress);

    if (!foundFiles.has(optionalFile)) {
      // Special case for icons directory - we might have individual icon files
      if (optionalFile === "icons" && foundDirs.has("icons")) {
        console.log("Found icons directory - acceptable");
      } else {
        warnings.push(`Missing optional file or directory: ${optionalFile}`);
        console.log(`Missing optional file: ${optionalFile}`);
      }
    }
  });

  // Validate content of key JSON files
  console.log("Validating JSON file content...");
  reportProgress(80);

  // Get all JSON files, including those in subdirectories
  const jsonFiles = files.filter((file) => {
    const fileName = file.getName();
    return fileName.endsWith(".json") && !fileName.endsWith("/");
  });

  jsonFiles.forEach((file, index) => {
    const fileName = file.getName();
    console.log(
      `Validating JSON file ${index + 1}/${jsonFiles.length}: ${fileName}`,
    );

    // Update progress (80-95%)
    const jsonProgress = 80 + Math.round((index / jsonFiles.length) * 15);
    reportProgress(jsonProgress);

    try {
      const content = file.getDataAsString();
      JSON.parse(content); // Just check if it's valid JSON
      console.log(`Validated JSON file: ${fileName}`);
    } catch (error) {
      // For nested files, include the full path in the error message
      const errorFileName = fileName.includes("/") ? fileName : fileName;
      errors.push(`Invalid JSON in file: ${errorFileName}`);
      console.warn(`Invalid JSON in file: ${errorFileName}`, error);
    }
  });

  // If we have no required files at all, but we have some files, try to be lenient
  if (errors.length === REQUIRED_FILES.length && files.length > 0) {
    console.log(
      "No required files found, but we have some files. Attempting to proceed anyway.",
    );
    errors.length = 0; // Clear errors
    warnings.push(
      "No standard configuration files found. Attempting to process anyway.",
    );
  }

  // If we have a package.json file but no metadata.json, we can use package.json for metadata
  if (!foundFiles.has("metadata.json") && foundFiles.has("package.json")) {
    console.log(
      "Found package.json but no metadata.json - will use package.json for metadata",
    );
    warnings.push("Using package.json for metadata information");
  }

  reportProgress(100);
  console.log(
    "Validation completed with " +
      (errors.length > 0 ? errors.length + " errors" : "no errors") +
      " and " +
      (warnings.length > 0 ? warnings.length + " warnings" : "no warnings"),
  );

  return {
    success: errors.length === 0,
    errors: errors.length > 0 ? errors : undefined,
    warnings: warnings.length > 0 ? warnings : undefined,
  };
}

// cleanupTempResources function moved to src/cleanup.ts
