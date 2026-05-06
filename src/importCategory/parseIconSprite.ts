/**
 * Functions for parsing and processing SVG sprite files (icons.svg)
 */

/**
 * Result of SVG sprite processing with error tracking
 */
interface SpriteProcessingResult {
  /** Successfully extracted icons */
  icons: { name: string; svg: string; id: string }[];

  /** Errors encountered during processing */
  errors: Array<{
    symbolId: string;
    error: string;
    context?: Record<string, any>;
  }>;
}

/**
 * Safe debug logger that falls back to Logger if debugLog is not available
 */
function safeSpriteDebugLog(message: string, forceFlush = false) {
  // Try debug logger first
  try {
    if (typeof debugLog === "function") {
      debugLog(message, forceFlush);
      return;
    }
  } catch (e) {
    // Fall through to Logger
  }

  // Fall back to standard logger
  try {
    console.log(message);
  } catch (e) {
    // Last resort
    console.log(message);
  }
}

/**
 * Processes an SVG sprite blob and extracts individual icons
 * @param tempFolder - The temporary folder to save the extracted icons
 * @param iconsSvgBlob - Optional blob containing the icons.svg file
 * @returns Result object with icons and errors
 */
function processIconSpriteBlob(
  tempFolder: GoogleAppsScript.Drive.Folder,
  iconsSvgBlob?: GoogleAppsScript.Base.Blob,
): SpriteProcessingResult {
  const result: SpriteProcessingResult = {
    icons: [],
    errors: []
  };

  try {
    console.log("Processing icons.svg sprite blob");

    // Get permanent config folder for icon storage
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(spreadsheet.getId())
      .getParents()
      .next();
    const configFolderName = slugify(spreadsheet.getName());

    console.log(`Looking for config folder: ${configFolderName}`);

    // Find or create config folder
    let configFolder: GoogleAppsScript.Drive.Folder;
    const configFolders = parentFolder.getFoldersByName(configFolderName);
    if (configFolders.hasNext()) {
      configFolder = configFolders.next();
      console.log(`Using existing config folder: ${configFolder.getName()}`);
    } else {
      configFolder = parentFolder.createFolder(configFolderName);
      console.log(`Created new config folder: ${configFolder.getName()}`);
    }

    // Find or create permanent icons folder
    let permanentIconsFolder: GoogleAppsScript.Drive.Folder;
    const iconsFolders = configFolder.getFoldersByName("icons");
    if (iconsFolders.hasNext()) {
      permanentIconsFolder = iconsFolders.next();
      console.log(`Using existing icons folder: ${permanentIconsFolder.getName()}`);
    } else {
      permanentIconsFolder = configFolder.createFolder("icons");
      console.log(`Created new icons folder: ${permanentIconsFolder.getName()}`);
    }

    // Create temporary folders for extraction
    const iconsSvgFolder = tempFolder.createFolder("icons_svg_temp");
    const iconsOutputFolder = tempFolder.createFolder("icons_output");

    // If blob provided, save it to the iconsSvgFolder
    if (iconsSvgBlob) {
      console.log("Saving provided icons.svg blob to temp folder");
      iconsSvgFolder.createFile(iconsSvgBlob);
    } else {
      console.log("No blob provided, checking for icons.svg in temp folder files");
    }

    // Process the sprite
    const icons = deconstructSvgSprite(
      iconsSvgFolder.getId(),
      iconsOutputFolder.getId(),
      result.errors,  // Pass errors array for tracking
    );

    console.log(`Extracted ${icons.length} icons from sprite`);

    // Move icons to permanent folder and update URLs
    const permanentIcons = icons.map((icon) => {
      try {
        // Find the icons subfolder in the temp output folder
        const iconsSubfolders = iconsOutputFolder.getFoldersByName("icons");
        if (!iconsSubfolders.hasNext()) {
          console.warn(`No icons subfolder found for ${icon.name}`);
          return icon;
        }

        const iconsSubfolder = iconsSubfolders.next();
        const fileName = `${icon.name}.svg`;
        const tempFiles = iconsSubfolder.getFilesByName(fileName);

        if (!tempFiles.hasNext()) {
          console.warn(`Icon file not found: ${fileName}`);
          return icon;
        }

        const tempFile = tempFiles.next();
        console.log(`Moving icon ${fileName} to permanent folder`);

        // Check if file already exists in permanent folder
        const existingFiles = permanentIconsFolder.getFilesByName(fileName);
        let permanentFile;

        if (existingFiles.hasNext()) {
          // Update existing file content
          permanentFile = existingFiles.next();
          permanentFile.setContent(tempFile.getBlob().getDataAsString());
          console.log(`Updated existing icon: ${fileName}`);
        } else {
          // Create new file in permanent folder
          permanentFile = permanentIconsFolder.createFile(
            tempFile.getBlob().setName(fileName),
          );
          console.log(`Created new icon: ${fileName}`);
        }

        // Return icon with permanent URL
        return {
          name: icon.name,
          svg: permanentFile.getUrl(),
          id: icon.id,
        };
      } catch (error) {
        console.error(`Error moving icon ${icon.name} to permanent folder:`, error);
        return icon;
      }
    });

    console.log(`Successfully moved ${permanentIcons.length} icons to permanent folder`);
    result.icons = permanentIcons;
    return result;
  } catch (error) {
    console.error("Error processing icon sprite blob:", error);
    result.errors.push({
      symbolId: "sprite-processing",
      error: `Failed to process SVG sprite: ${error.message}`,
      context: { stack: error.stack }
    });
    return result;
  }
}

/**
 * Checks if a file with the given name exists in a folder
 * @param folder - The folder to check
 * @param fileName - The name of the file to look for
 * @returns True if the file exists, false otherwise
 */
function fileExistsInFolder(
  folder: GoogleAppsScript.Drive.Folder,
  fileName: string,
): boolean {
  const files = folder.getFilesByName(fileName);
  return files.hasNext();
}

/**
 * Deconstructs an SVG sprite (icons.svg) stored in a Google Drive folder into separate SVG icons as files.
 * @param configFolderId - Google Drive folder ID containing icons.svg
 * @param outputFolderId - Google Drive folder ID to save the individual icons into a subfolder called 'icons'
 * @param errors - Optional array to collect errors during processing
 * @returns An array of icon objects with name and URL
 */
function deconstructSvgSprite(
  configFolderId: string,
  outputFolderId: string,
  errors?: Array<{ symbolId: string; error: string; context?: Record<string, any> }>,
): { name: string; svg: string; id: string }[] {
  try {
    safeSpriteDebugLog(
      `Deconstructing SVG sprite from ${configFolderId} to ${outputFolderId}`,
    );

    const configFolder = DriveApp.getFolderById(configFolderId);
    const outputFolder = DriveApp.getFolderById(outputFolderId);

    // Check if icons.svg exists
    const files = configFolder.getFilesByName("icons.svg");
    if (!files.hasNext()) {
      safeSpriteDebugLog(
        "No icons.svg found in folder, skipping SVG sprite deconstruction",
      );
      return [];
    }

    const iconsFile = files.next();
    const fileBlob = iconsFile.getBlob();
    safeSpriteDebugLog(
      `Processing SVG file: ${iconsFile.getName()} (${fileBlob.getContentType()})`,
    );

    const svgContent = fileBlob.getDataAsString();
    safeSpriteDebugLog(`SVG content length: ${svgContent.length} characters`);
    safeSpriteDebugLog(`SVG content preview: ${svgContent.substring(0, 200)}...`);

    const xml = XmlService.parse(svgContent);
    const svgRoot = xml.getRootElement();
    safeSpriteDebugLog(
      `SVG root element name: ${svgRoot.getName()}, namespace: ${svgRoot.getNamespace().getURI()}`,
    );

    // Get all children to see what's available
    const allChildren = svgRoot.getChildren();
    safeSpriteDebugLog(
      `SVG root has ${allChildren.length} total children of various types`,
    );
    safeSpriteDebugLog(
      `SVG root children types: ${allChildren.map((child) => child.getName()).join(", ")}`,
    );

    // Get symbol elements
    let symbols = svgRoot.getChildren("symbol");
    if (!symbols || symbols.length === 0) {
      // Alternative approach: filter all children to find symbols
      symbols = svgRoot
        .getChildren()
        .filter((child) => child.getName() === "symbol");
      safeSpriteDebugLog(
        `Found ${symbols.length} symbols using alternative filter approach`,
      );
    }

    safeSpriteDebugLog(`Found ${symbols.length} symbols in SVG sprite`);

    if (!symbols || symbols.length === 0) {
      const errorMsg = "No <symbol> children in SVG";
      safeSpriteDebugLog(errorMsg);
      if (errors) {
        errors.push({
          symbolId: "sprite",
          error: errorMsg,
          context: {
            rootChildren: allChildren.map((c) => c.getName()),
            totalChildren: allChildren.length
          }
        });
      }
      return [];
    }

    // Log first few symbols if they exist
    if (symbols.length > 0) {
      const sampleSymbol = symbols[0];
      safeSpriteDebugLog(
        `First symbol ID: ${sampleSymbol.getAttribute("id")?.getValue() || "no id"}`,
      );
      safeSpriteDebugLog(
        `First symbol children count: ${sampleSymbol.getChildren().length}`,
      );
      safeSpriteDebugLog(
        `First symbol children types: ${sampleSymbol
          .getChildren()
          .map((c) => c.getName())
          .join(", ")}`,
      );
    }

    // Create / find 'icons' subfolder in output folder
    let iconsSubfolder: GoogleAppsScript.Drive.Folder;
    const subfolders = outputFolder.getFoldersByName("icons");
    if (subfolders.hasNext()) {
      iconsSubfolder = subfolders.next();
      safeSpriteDebugLog(`Using existing icons subfolder: ${iconsSubfolder.getName()}`);
    } else {
      iconsSubfolder = outputFolder.createFolder("icons");
      safeSpriteDebugLog(`Created new icons subfolder: ${iconsSubfolder.getName()}`);
    }

    // Process each icon
    const iconObjects: { name: string; svg: string; id: string }[] = [];

    safeSpriteDebugLog(`\n=== SVG ICON EXTRACTION (${symbols.length} symbols) ===`);

    for (let i = 0; i < symbols.length; i++) {
      try {
        const symbol = symbols[i];
        const idAttr = symbol.getAttribute("id");

        if (!idAttr) {
          const symbolId = `symbol-${i + 1}`;
          safeSpriteDebugLog(`⚠️  Symbol #${i + 1}: No id attribute, skipping`);
          if (errors) {
            errors.push({
              symbolId,
              error: "Missing id attribute",
              context: { symbolIndex: i }
            });
          }
          continue;
        }

        const id = idAttr.getValue();

        safeSpriteDebugLog(`Processing symbol #${i + 1}: id=${id}`);
        safeSpriteDebugLog(`Symbol children count: ${symbol.getChildren().length}`);
        safeSpriteDebugLog(
          `Symbol children types: ${symbol
            .getChildren()
            .map((c) => c.getName())
            .join(", ")}`,
        );

        // Build new SVG root with a single symbol's content
        // CRITICAL: Create SVG element with proper namespace
        const svgNamespace = svgRoot.getNamespace();
        const newSvg = XmlService.createElement("svg", svgNamespace);
        safeSpriteDebugLog(`Created new SVG with namespace: ${svgNamespace.getURI()}`);

        // Copy all attributes from root <svg> to new <svg>
        const attrs = svgRoot.getAttributes();
        for (let k = 0; k < attrs.length; k++) {
          const attr = attrs[k];
          newSvg.setAttribute(attr.getName(), attr.getValue());
        }

        // Use symbol's viewBox if available, otherwise use root or default
        const symbolViewBox = symbol.getAttribute("viewBox");
        if (symbolViewBox) {
          newSvg.setAttribute("viewBox", symbolViewBox.getValue());
          safeSpriteDebugLog(`Using symbol's viewBox: ${symbolViewBox.getValue()}`);
        } else if (svgRoot.getAttribute("viewBox")) {
          newSvg.setAttribute("viewBox", svgRoot.getAttribute("viewBox").getValue());
          safeSpriteDebugLog(`Using root SVG viewBox: ${svgRoot.getAttribute("viewBox").getValue()}`);
        } else {
          newSvg.setAttribute("viewBox", "0 0 24 24");
          safeSpriteDebugLog(`Using default viewBox: 0 0 24 24`);
        }

        // Get ALL child elements from the symbol (path, rect, circle, etc.)
        const children = symbol.getChildren();
        safeSpriteDebugLog(`Symbol ${id} has ${children.length} child element(s)`);

        // Clone each child element with all its attributes and content
        for (let c = 0; c < children.length; c++) {
          const child = children[c];
          const childName = child.getName();
          safeSpriteDebugLog(`  - Processing child element #${c + 1}: <${childName}>`);

          // Create new element with same name and namespace
          const newChild = XmlService.createElement(childName, child.getNamespace());

          // Copy all attributes
          const childAttrs = child.getAttributes();
          safeSpriteDebugLog(`    - Has ${childAttrs.length} attribute(s)`);
          for (let a = 0; a < childAttrs.length; a++) {
            const attr = childAttrs[a];
            newChild.setAttribute(attr.getName(), attr.getValue());
            safeSpriteDebugLog(`      - ${attr.getName()}="${attr.getValue()}"`);
          }

          // Copy text content if any
          const textContent = child.getText();
          if (textContent) {
            newChild.setText(textContent);
          }

          // Recursively copy nested children if any
          const nestedChildren = child.getChildren();
          if (nestedChildren.length > 0) {
            safeSpriteDebugLog(`    - Has ${nestedChildren.length} nested child(ren)`);
            // For nested children, we'll use detach since they're less common
            for (let n = 0; n < nestedChildren.length; n++) {
              newChild.addContent(nestedChildren[n].detach());
            }
          }

          // Add the cloned child to the new SVG
          newSvg.addContent(newChild);
        }

        const newDoc = XmlService.createDocument(newSvg);
        const newXmlString = XmlService.getPrettyFormat().format(newDoc);

        // Debug log to check what's being created
        safeSpriteDebugLog(`Generated SVG XML for ${id}:`);
        if (newXmlString.length <= 500) {
          safeSpriteDebugLog(newXmlString); // Log full content if small
        } else {
          safeSpriteDebugLog(newXmlString.substring(0, 500) + "..."); // Log first 500 chars
        }

        // Verify namespace is present in output
        if (newXmlString.indexOf('xmlns=') !== -1) {
          safeSpriteDebugLog(`✓ Namespace present in generated SVG`);
        } else {
          safeSpriteDebugLog(`✗ WARNING: Namespace missing in generated SVG!`);
          if (errors) {
            errors.push({
              symbolId: id,
              error: "Missing namespace in generated SVG",
              context: { symbolIndex: i }
            });
          }
          continue;
        }

        // Create a single file with format {id}.svg using full symbol ID
        const fileName = `${id}.svg`;

        // Write the file (overwrite if exists since we're using full IDs now)
        const iconFile = iconsSubfolder.createFile(
          fileName,
          newXmlString,
          MimeType.SVG,
        );
        safeSpriteDebugLog(`✓ Wrote ${fileName} successfully`);

        // Add to icon objects using full ID
        iconObjects.push({
          name: id,
          svg: iconFile.getUrl(),
          id: id,
        });
      } catch (error) {
        // Isolate failures to prevent one bad symbol from breaking all icons
        const symbolId = i < symbols.length && symbols[i].getAttribute("id")
          ? symbols[i].getAttribute("id").getValue()
          : `symbol-${i + 1}`;
        safeSpriteDebugLog(`❌ ERROR processing symbol #${i + 1} (${symbolId}): ${error}`);
        safeSpriteDebugLog(`   Stack: ${error.stack || "No stack trace"}`);
        console.error(`Failed to process symbol ${symbolId}:`, error);
        if (errors) {
          errors.push({
            symbolId,
            error: String(error),
            context: { symbolIndex: i, stack: error.stack }
          });
        }
        // Continue processing other symbols
        continue;
      }
    }

    // Report extraction results
    safeSpriteDebugLog(`\n=== SVG EXTRACTION RESULTS ===`);
    safeSpriteDebugLog(`✓ Successfully extracted: ${iconObjects.length}/${symbols.length} icons`);

    if (errors && errors.length > 0) {
      safeSpriteDebugLog(`❌ Failed to extract: ${errors.length}/${symbols.length} icons`);
      safeSpriteDebugLog(`\n=== FAILED SYMBOLS ===`);
      errors.forEach((failed, index) => {
        safeSpriteDebugLog(`  ${index + 1}. Symbol "${failed.symbolId}": ${failed.error}`);
      });
      safeSpriteDebugLog(`These icons will fall back to PNG extraction if available.`);
      safeSpriteDebugLog(`=== END FAILED SYMBOLS ===`);
    }
    safeSpriteDebugLog(`=== END SVG EXTRACTION ===\n`);

    return iconObjects;
  } catch (error) {
    safeSpriteDebugLog(`Error in deconstructSvgSprite: ${error}`);
    console.error(`SVG sprite processing error: ${error}`);
    return [];
  }
}
