/**
 * Creates a ZIP archive from a Google Drive folder
 *
 * Recursively collects all files from the folder and its subfolders,
 * then creates a ZIP blob. Includes retry logic for Drive API reliability
 * and progress reporting for long operations.
 *
 * @param folderId - Google Drive folder ID to archive
 * @param onProgress - Optional callback for progress updates
 * @returns ZIP file as a Blob
 * @throws Error if folder cannot be accessed or depth limit exceeded
 *
 * @example
 * const zip = saveDriveFolderToZip(folderId, (msg, detail) => {
 *   console.log(`${msg}: ${detail}`);
 * });
 * // Returns: Blob containing ZIP archive of folder contents
 */
const ZIP_DIGEST_LOGGING_PROPERTY_KEY = "ENABLE_ZIP_DIGEST_LOGGING";
const MAX_DRIVE_FOLDER_DEPTH = 50;
const ICON_HASH_PROPERTY_PREFIX = "ICON_HASH_V1_";
const SHARING_POLICY_PROPERTY_KEY = "DEFAULT_SHARING_POLICY";
const DEFAULT_SHARING_POLICY: SharingPolicy = "anyone";

interface IconHashMetadata {
  hash: string;
  fileId?: string;
  fileUrl?: string;
  updatedAt?: string;
}

type MaybeDriveFolder = GoogleAppsScript.Drive.Folder | null;

interface SaveConfigOptions {
  skipDriveWrites?: boolean;
}

interface ConfigFolders {
  presets: MaybeDriveFolder;
  icons: MaybeDriveFolder;
  fields: MaybeDriveFolder;
  messages: MaybeDriveFolder;
}

interface IconCacheStats {
  reusedFromCache: number;
  driveWrites: number;
  inMemoryUrls: number;
  generatedFallbacks: number;
}

type SharingPolicy = "inherit" | "anyone" | "off";

let iconHashCache: Record<string, IconHashMetadata> | null = null;
let iconHashKeysUsedThisRun: Record<string, boolean> = {};
let cachedSharingPolicy: SharingPolicy | null = null;

function saveDriveFolderToZip(
  folderId,
  onProgress?: (message: string, detail?: string) => void,
  preCollectedBlobs?: GoogleAppsScript.Base.Blob[],
): GoogleAppsScript.Base.Blob {
  getScopedLogger("DriveService").info("[ZIP] Attempting to access folder with ID:", folderId);

  if (!folderId) {
    throw new Error("Folder ID is null or undefined");
  }

  // Progress callback helper
  const reportProgress = (message: string, detail?: string) => {
    getScopedLogger("DriveService").info(`[ZIP] ${message}${detail ? ': ' + detail : ''}`);
    if (onProgress) {
      onProgress(message, detail);
    }
  };

  let folder: GoogleAppsScript.Drive.Folder;
  let retryCount = 0;
  const maxRetries = 3;

  // Retry logic for accessing folder (handles transient Drive API issues)
  while (retryCount <= maxRetries) {
    try {
      folder = DriveApp.getFolderById(folderId);
      break; // Success, exit retry loop
    } catch (error) {
      retryCount++;
      getScopedLogger("DriveService").error(`[ZIP] Failed to access folder (attempt ${retryCount}/${maxRetries + 1}):`, error.message);

      if (retryCount > maxRetries) {
        throw new Error(`Failed to access Drive folder with ID "${folderId}" after ${maxRetries + 1} attempts. This could be due to permissions or the folder not being properly created. Original error: ${error.message}`);
      }

      // Wait before retrying (exponential backoff)
      const waitTime = 1000 * Math.pow(2, retryCount);
      getScopedLogger("DriveService").info(`[ZIP] Waiting ${waitTime}ms before retry...`);
      Utilities.sleep(waitTime);
    }
  }

  if (!folder) {
    throw new Error(`Folder with ID "${folderId}" was not found or is not accessible`);
  }

  getScopedLogger("DriveService").info("[ZIP] Successfully accessed folder:", folder.getName());
  const logZipDigests = isZipDigestLoggingEnabled();

  if (preCollectedBlobs && preCollectedBlobs.length > 0) {
    reportProgress("Creating package... (5/8)", `Using cached blobs (${preCollectedBlobs.length})...`);
    const zipStartTime = new Date().getTime();
    const sortedCachedBlobs = sortBlobsByName(preCollectedBlobs);
    const zipBlob = Utilities.zip(sortedCachedBlobs, `${folder.getName()}.zip`);
    const zipTime = ((new Date().getTime() - zipStartTime) / 1000).toFixed(1);
    getScopedLogger("DriveService").info(`[ZIP] ZIP archive created from cached blobs in ${zipTime}s (size: ${zipBlob.getBytes().length} bytes)`);
    if (logZipDigests) {
      logZipDigestComparison(sortedCachedBlobs, zipBlob, folder);
    }
    return zipBlob;
  }

  reportProgress("Creating package... (5/8)", "Collecting files...");

  const blobs: GoogleAppsScript.Base.Blob[] = [];
  let fileCount = 0;
  const startTime = new Date().getTime();

  function addFolderContentsToBlobs(
    currentFolder: GoogleAppsScript.Drive.Folder,
    path = "",
    depth = 0,
  ) {
    // Prevent infinite recursion with depth limit
    if (depth > MAX_DRIVE_FOLDER_DEPTH) {
      getScopedLogger("DriveService").error(`[ZIP] Maximum folder depth (${MAX_DRIVE_FOLDER_DEPTH}) exceeded at path: ${path}`);
      throw new Error(`Maximum folder depth of ${MAX_DRIVE_FOLDER_DEPTH} exceeded. Please check for circular folder references or deeply nested structures.`);
    }

    const files = currentFolder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      getScopedLogger("DriveService").info(`[ZIP] Adding file ${fileCount}: ${path}${file.getName()}`);

      // Heartbeat logging and progress reporting every 10 files
      if (fileCount % 10 === 0) {
        const elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(1);
        getScopedLogger("DriveService").info(`[ZIP] Progress: ${fileCount} files processed in ${elapsed}s`);
        reportProgress("Creating package... (5/8)", `Collecting files (${fileCount})...`);
      }

      blobs.push(file.getBlob().setName(`${path}${file.getName()}`));
    }

    const subFolders = currentFolder.getFolders();
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      const newPath = `${path}${subFolder.getName()}/`;
      addFolderContentsToBlobs(subFolder, newPath, depth + 1);
    }
  }

  addFolderContentsToBlobs(folder);

  const totalTime = ((new Date().getTime() - startTime) / 1000).toFixed(1);
  getScopedLogger("DriveService").info(`[ZIP] Collected ${fileCount} files in ${totalTime}s`);
  reportProgress("Creating package... (5/8)", `Compressing ${fileCount} files into archive...`);

  const zipStartTime = new Date().getTime();
  const sortedDriveBlobs = sortBlobsByName(blobs);
  const zipBlob = Utilities.zip(sortedDriveBlobs, `${folder.getName()}.zip`);
  const zipTime = ((new Date().getTime() - zipStartTime) / 1000).toFixed(1);

  getScopedLogger("DriveService").info(`[ZIP] ZIP archive created in ${zipTime}s (size: ${zipBlob.getBytes().length} bytes)`);
  if (logZipDigests) {
    const driveDigest = computeBlobMd5Digest(zipBlob);
    getScopedLogger("DriveService").info(`[ZIP] Digest (Drive traversal): ${driveDigest}`);
  }
  return zipBlob;
}

function isZipDigestLoggingEnabled(): boolean {
  try {
    const propertyValue = PropertiesService.getScriptProperties().getProperty(
      ZIP_DIGEST_LOGGING_PROPERTY_KEY,
    );
    const isEnabled = typeof propertyValue === "string" && propertyValue.toLowerCase() === "true";
    if (isEnabled) {
      getScopedLogger("DriveService").info("[ZIP] Digest comparison logging enabled via script property");
    }
    return isEnabled;
  } catch (error) {
    getScopedLogger("DriveService").warn(
      `[ZIP] Unable to read script property "${ZIP_DIGEST_LOGGING_PROPERTY_KEY}": ${error.message}`,
    );
    return false;
  }
}

function logZipDigestComparison(
  cachedBlobs: GoogleAppsScript.Base.Blob[],
  cachedZipBlob: GoogleAppsScript.Base.Blob,
  sourceFolder: GoogleAppsScript.Drive.Folder,
): void {
  try {
    const cachedNames = cachedBlobs.map((blob) => blob.getName() || "");
    const cachedSortedNames = cachedNames.slice().sort();
    const cachedDigest = computeBlobMd5Digest(cachedZipBlob);
    getScopedLogger("DriveService").info(`[ZIP] Digest (cached blobs): ${cachedDigest}`);

    const driveBlobs = collectFolderBlobsForDigest(sourceFolder);
    const driveNames = driveBlobs.map((blob) => blob.getName() || "");
    const driveSortedNames = driveNames.slice().sort();
    logZipBlobNameDiagnostics(cachedSortedNames, driveSortedNames);

    const sortedDriveBlobs = sortBlobsByName(driveBlobs);
    const driveZipBlob = Utilities.zip(sortedDriveBlobs, `${sourceFolder.getName()}.zip`);
    const driveDigest = computeBlobMd5Digest(driveZipBlob);
    getScopedLogger("DriveService").info(`[ZIP] Digest (Drive traversal): ${driveDigest}`);

    logPerBlobDigestDifferences(cachedBlobs, sortedDriveBlobs);

    const digestsMatch = cachedDigest === driveDigest;
    getScopedLogger("DriveService").info(`[ZIP] Digest comparison result: ${digestsMatch ? "MATCH" : "MISMATCH"}`);
  } catch (error) {
    getScopedLogger("DriveService").error("[ZIP] Failed to compute ZIP digest comparison:", error);
  }
}

function logZipBlobNameDiagnostics(
  cachedSortedNames: string[],
  driveSortedNames: string[],
): void {
  const cachedCount = cachedSortedNames.length;
  const driveCount = driveSortedNames.length;

  if (
    cachedCount === driveCount &&
    cachedSortedNames.every((name, index) => name === driveSortedNames[index])
  ) {
    getScopedLogger("DriveService").info(`[ZIP] Blob name lists match (${cachedCount} entries)`);
  } else {
    getScopedLogger("DriveService").info(
      `[ZIP] Blob name lists differ — cached (${cachedCount}) vs Drive (${driveCount})`,
    );
    getScopedLogger("DriveService").info(
      `[ZIP] Cached blob names: ${cachedSortedNames.join(", ") || "(none)"}`,
    );
    getScopedLogger("DriveService").info(
      `[ZIP] Drive blob names: ${driveSortedNames.join(", ") || "(none)"}`,
    );

    const cachedOnly = computeNameDifference(
      cachedSortedNames,
      new Set(driveSortedNames),
    );
    const driveOnly = computeNameDifference(
      driveSortedNames,
      new Set(cachedSortedNames),
    );

    getScopedLogger("DriveService").info(
      `[ZIP] Only in cached blobs: ${cachedOnly.join(", ") || "(none)"}`,
    );
    getScopedLogger("DriveService").info(
      `[ZIP] Only in Drive blobs: ${driveOnly.join(", ") || "(none)"}`,
    );
  }

  const cachedDuplicates = findDuplicateNames(cachedSortedNames);
  const driveDuplicates = findDuplicateNames(driveSortedNames);

  if (cachedDuplicates.length > 0) {
    getScopedLogger("DriveService").info(
      `[ZIP] Duplicate cached blob names: ${cachedDuplicates.join(", ")}`,
    );
  }
  if (driveDuplicates.length > 0) {
    getScopedLogger("DriveService").info(
      `[ZIP] Duplicate Drive blob names: ${driveDuplicates.join(", ")}`,
    );
  }
}

function computeNameDifference(
  names: string[],
  comparisonSet: Set<string>,
): string[] {
  const difference = new Set<string>();
  for (const name of names) {
    if (!comparisonSet.has(name)) {
      difference.add(name);
    }
  }
  return Array.from(difference.values()).sort();
}

function findDuplicateNames(names: string[]): string[] {
  const duplicates = new Set<string>();
  const counts: Record<string, number> = {};
  for (const name of names) {
    counts[name] = (counts[name] || 0) + 1;
    if (counts[name] === 2) {
      duplicates.add(name);
    }
  }
  return Array.from(duplicates.values()).sort();
}

function logPerBlobDigestDifferences(
  cachedBlobs: GoogleAppsScript.Base.Blob[],
  driveBlobs: GoogleAppsScript.Base.Blob[],
): void {
  const cachedDigestMap = computeBlobDigestMap(cachedBlobs);
  const driveDigestMap = computeBlobDigestMap(driveBlobs);
  const mismatchedNames: string[] = [];

  Object.keys(cachedDigestMap).forEach((name) => {
    if (cachedDigestMap[name] !== driveDigestMap[name]) {
      mismatchedNames.push(name);
    }
  });

  if (mismatchedNames.length === 0) {
    getScopedLogger("DriveService").info("[ZIP] Individual blob digests match for all entries");
  } else {
    getScopedLogger("DriveService").info(
      `[ZIP] Individual blob digest mismatches (${mismatchedNames.length}): ${mismatchedNames.join(", ")}`,
    );
    mismatchedNames.forEach((name) => {
      getScopedLogger("DriveService").info(
        `[ZIP]  • ${name}: cached=${cachedDigestMap[name]} vs Drive=${driveDigestMap[name]}`,
      );
    });
  }
}

function computeBlobMd5Digest(blob: GoogleAppsScript.Base.Blob): string {
  const digestBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    blob.getBytes(),
  );
  return digestBytes
    .map((byte) => {
      const normalized = byte < 0 ? byte + 256 : byte;
      return normalized.toString(16).padStart(2, "0");
    })
    .join("");
}

function computeBlobDigestMap(
  blobs: GoogleAppsScript.Base.Blob[],
): Record<string, string> {
  const digestMap: Record<string, string> = {};
  blobs.forEach((blob) => {
    const name = blob.getName() || "";
    digestMap[name] = computeBlobMd5Digest(blob);
  });
  return digestMap;
}

function sortBlobsByName(
  blobs: GoogleAppsScript.Base.Blob[],
): GoogleAppsScript.Base.Blob[] {
  return blobs
    .map((blob) => {
      const name = blob.getName() || "";
      const copy = blob.copyBlob();
      if (copy.getName() !== name) {
        copy.setName(name);
      }
      return {
        name,
        blob: copy,
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name))
    .map((entry) => entry.blob);
}

function collectFolderBlobsForDigest(
  folder: GoogleAppsScript.Drive.Folder,
  path = "",
  depth = 0,
  blobs: GoogleAppsScript.Base.Blob[] = [],
): GoogleAppsScript.Base.Blob[] {
  if (depth > MAX_DRIVE_FOLDER_DEPTH) {
    throw new Error(
      `Maximum folder depth of ${MAX_DRIVE_FOLDER_DEPTH} exceeded while collecting digest comparison blobs.`,
    );
  }

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    blobs.push(file.getBlob().copyBlob().setName(`${path}${file.getName()}`));
  }

  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    const nextPath = `${path}${subFolder.getName()}/`;
    collectFolderBlobsForDigest(subFolder, nextPath, depth + 1, blobs);
  }

  return blobs;
}

function getConfigFolder(): GoogleAppsScript.Drive.Folder {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const configFolderName = slugify(spreadsheet.getName());

  let configFolder: GoogleAppsScript.Drive.Folder;
  const configFolders =
    DriveApp.getRootFolder().getFoldersByName(configFolderName);

  if (configFolders.hasNext()) {
    configFolder = configFolders.next();
  } else {
    configFolder = DriveApp.getRootFolder().createFolder(configFolderName);
  }
  applyDefaultSharing(configFolder);
  return configFolder;
}

function saveZipToDrive(zipBlob: GoogleAppsScript.Base.Blob, version): string {
  getScopedLogger("DriveService").info("Saving ZIP file to Drive...");
  const configFolder = getConfigFolder();
  const buildsFolder = "builds";
  let buildsFolderObj: GoogleAppsScript.Drive.Folder;
  const buildsFolders = configFolder.getFoldersByName(buildsFolder);
  if (buildsFolders.hasNext()) {
    buildsFolderObj = buildsFolders.next();
    applyDefaultSharing(buildsFolderObj);
  } else {
    buildsFolderObj = configFolder.createFolder(buildsFolder);
    applyDefaultSharing(buildsFolderObj);
  }
  getScopedLogger("DriveService").info("Saving ZIP file to Drive...");
  const fileName = `${version}.zip`;
  const doubleZippedBlob = Utilities.zip([zipBlob], fileName);
  const zipFile = buildsFolderObj
    .createFile(doubleZippedBlob)
    .setName(fileName);
  applyDefaultSharing(zipFile);
  const fileUrl = zipFile.getUrl();
  getScopedLogger("DriveService").info(`Download the ZIP file here: ${fileUrl}`);

  return fileUrl;
}

/**
 * Saves CoMapeo configuration to Google Drive
 *
 * Creates a structured folder hierarchy in Drive and saves all configuration
 * files including fields.json, presets.json, icons, and translations.
 * Includes retry logic and progress reporting.
 *
 * @param config - Complete CoMapeo configuration object
 * @param onProgress - Optional callback for progress updates
 * @returns Object with folder URL and ID
 * @throws Error if folder creation fails or save operation fails
 *
 * @example
 * const result = saveConfigToDrive(config, (msg, detail) => {
 *   console.log(`${msg}: ${detail}`);
 * });
 * // Returns: { url: "https://drive.google.com/...", id: "abc123" }
 */
function saveConfigToDrive(
  config: CoMapeoConfig,
  onProgress?: (message: string, detail?: string) => void,
  options?: SaveConfigOptions,
): { url: string; id: string; zipBlobs: GoogleAppsScript.Base.Blob[] } {
  const skipDriveWrites = Boolean(options?.skipDriveWrites);
  const shouldWriteToDrive = !skipDriveWrites;
  const folderName = `${slugify(config.metadata.version)}`;
  getScopedLogger("DriveService").info(
    `[DRIVE] Saving config to ${shouldWriteToDrive ? "drive" : "in-memory zip"}:`,
    folderName,
  );
  const startTime = new Date().getTime();
  const zipBlobs: GoogleAppsScript.Base.Blob[] = [];
  resetIconHashTracking();

  // Progress callback helper
  const reportProgress = (message: string, detail?: string) => {
    getScopedLogger("DriveService").info(`[DRIVE] ${message}${detail ? ': ' + detail : ''}`);
    if (onProgress) {
      onProgress(message, detail);
    }
  };

  const configFolder: GoogleAppsScript.Drive.Folder = getConfigFolder();

  let rootFolder: GoogleAppsScript.Drive.Folder | null = null;
  let folderId = "";
  let folderUrl = "";

  let retryCount = 0;
  const maxRetries = 3;

  // Retry logic for folder creation (handles Drive API rate limits)
  while (retryCount <= maxRetries) {
    try {
      const rawBuildsFolder = getOrCreateSubFolder(configFolder, "rawBuilds");
      applyDefaultSharing(rawBuildsFolder);
      rootFolder = getOrCreateSubFolder(rawBuildsFolder, folderName, true);
      applyDefaultSharing(rootFolder);
      break; // Success, exit retry loop
    } catch (error) {
      retryCount++;
      getScopedLogger("DriveService").error(`[DRIVE] Error creating folder (attempt ${retryCount}/${maxRetries + 1}):`, error.message);

      if (retryCount > maxRetries) {
        throw new Error(
          `Failed to create folder "${folderName}" in "rawBuilds" after ${maxRetries + 1} attempts. Please check your Drive permissions and quota, then try again. Original error: ${error.message}`,
        );
      }

      // Wait before retrying (exponential backoff)
      const waitTime = 1000 * Math.pow(2, retryCount);
      getScopedLogger("DriveService").info(`[DRIVE] Waiting ${waitTime}ms before retry...`);
      Utilities.sleep(waitTime);
    }
  }

  if (!rootFolder) {
    throw new Error(
      `Failed to create folder "${folderName}" in "rawBuilds". Root folder is undefined.`,
    );
  }
  getScopedLogger("DriveService").info("[DRIVE] Created folder:", rootFolder.getName(), "in rawBuilds");

  // Verify folder ID is valid before proceeding
  folderId = rootFolder.getId();
  folderUrl = rootFolder.getUrl();
  getScopedLogger("DriveService").info("[DRIVE] Folder ID:", folderId);

  if (!folderId) {
    throw new Error("Failed to get valid folder ID from created folder");
  }

  if (!shouldWriteToDrive) {
    reportProgress("Saving to Drive... (4/8)", "Staging files in memory (Drive writes skipped – icons still updated)...");
  }

  try {
    reportProgress("Saving to Drive... (4/8)", "Creating folders...");
    const folders = createSubFolders(rootFolder, shouldWriteToDrive);
    getScopedLogger("DriveService").info("[DRIVE] ✅ Created subfolders successfully");

    // Process each step individually with better error handling and progress logging
    const iconSuffixes = ["-100px", "-24px"];

    reportProgress("Saving to Drive... (4/8)", `Saving ${config.presets.length} presets and icons...`);
    const presetsStart = new Date().getTime();
    savePresetsAndIcons(config, folders, iconSuffixes, zipBlobs, shouldWriteToDrive);
    const presetsTime = ((new Date().getTime() - presetsStart) / 1000).toFixed(1);
    getScopedLogger("DriveService").info(`[DRIVE] ✅ Saved presets and icons (${presetsTime}s)`);

    reportProgress("Saving to Drive... (4/8)", `Saving ${config.fields.length} fields...`);
    const fieldsStart = new Date().getTime();
    saveFields(config.fields, folders.fields, zipBlobs, shouldWriteToDrive);
    const fieldsTime = ((new Date().getTime() - fieldsStart) / 1000).toFixed(1);
    getScopedLogger("DriveService").info(`[DRIVE] ✅ Saved ${config.fields.length} fields (${fieldsTime}s)`);

    const languageCount = Object.keys(config.messages).length;
    reportProgress("Saving to Drive... (4/8)", `Saving translations for ${languageCount} languages...`);
    const messagesStart = new Date().getTime();
    saveMessages(config.messages, folders.messages, zipBlobs, shouldWriteToDrive);
    const messagesTime = ((new Date().getTime() - messagesStart) / 1000).toFixed(1);
    getScopedLogger("DriveService").info(`[DRIVE] ✅ Saved messages for ${languageCount} languages (${messagesTime}s)`);

    reportProgress("Saving to Drive... (4/8)", "Saving metadata and package...");
    const metadataStart = new Date().getTime();
    saveMetadataAndPackage(config, rootFolder, zipBlobs, shouldWriteToDrive);
    const metadataTime = ((new Date().getTime() - metadataStart) / 1000).toFixed(1);
    getScopedLogger("DriveService").info(`[DRIVE] ✅ Saved metadata and package (${metadataTime}s)`);

    const totalTime = ((new Date().getTime() - startTime) / 1000).toFixed(1);
    getScopedLogger("DriveService").info(`[DRIVE] ✅ Successfully saved all config files to folder (total: ${totalTime}s)`);

    const iconVariantCount = iconSuffixes.length > 0 ? iconSuffixes.length : 1;
    const iconFilesCreated = config.icons.length * iconVariantCount;
    const totalFilesCreated =
      config.presets.length +
      iconFilesCreated +
      config.fields.length +
      languageCount +
      2; // metadata.json + package.json

    if (shouldWriteToDrive) {
      if (totalFilesCreated > 80) {
        getScopedLogger("DriveService").info(`[DRIVE] Created ${totalFilesCreated} files; waiting for Drive sync...`);
        Utilities.sleep(2000);
      } else {
        getScopedLogger("DriveService").info("[DRIVE] Skipping Drive sync wait (small batch)");
      }
    }

    return {
      url: folderUrl,
      id: folderId,
      zipBlobs,
    };
  } catch (error) {
    getScopedLogger("DriveService").error("[DRIVE] Error saving config files to Drive:", error);
    throw new Error(`Failed to save config files to Drive: ${error.message}`);
  }
}

function createSubFolders(
  rootFolder: GoogleAppsScript.Drive.Folder,
  shouldWriteToDrive: boolean,
): ConfigFolders {
  const iconsFolder = getOrCreateSubFolder(rootFolder, "icons");
  applyDefaultSharing(iconsFolder);

  if (!shouldWriteToDrive) {
    return {
      presets: null,
      icons: iconsFolder,
      fields: null,
      messages: null,
    };
  }

  const presetsFolder = getOrCreateSubFolder(rootFolder, "presets");
  const fieldsFolder = getOrCreateSubFolder(rootFolder, "fields");
  const messagesFolder = getOrCreateSubFolder(rootFolder, "messages");

  applyDefaultSharing(presetsFolder);
  applyDefaultSharing(fieldsFolder);
  applyDefaultSharing(messagesFolder);

  return {
    presets: presetsFolder,
    icons: iconsFolder,
    fields: fieldsFolder,
    messages: messagesFolder,
  };
}

function savePresetsAndIcons(
  config: CoMapeoConfig,
  folders: ConfigFolders,
  suffixes: string[],
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  shouldWriteToDrive: boolean,
) {
  try {
    getScopedLogger("DriveService").info("Saving presets...");
    savePresets(config.presets, folders.presets, zipBlobs, shouldWriteToDrive);
    getScopedLogger("DriveService").info("Saving icons from cached config...");
    config.icons = saveExistingIconsToFolder(
      config,
      folders.icons,
      suffixes,
      zipBlobs,
      true,
    );
    getScopedLogger("DriveService").info("Icons saved successfully");
  } catch (error) {
    getScopedLogger("DriveService").error("Error in savePresetsAndIcons:", error);
    throw new Error(`Failed to save presets and icons: ${error.message}`);
  }
}

function saveExistingIconsToFolder(
  config: CoMapeoConfig,
  iconsFolder: MaybeDriveFolder,
  suffixes: string[],
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  shouldWriteToDrive: boolean,
): CoMapeoIcon[] {
  const iconMap = new Map<string, CoMapeoIcon>();
  for (const icon of config.icons) {
    const normalizedName = sanitizeIconSlug(icon.name) || icon.name;
    iconMap.set(normalizedName, { ...icon, name: normalizedName });
  }

  const { categories, backgroundColors, categoriesSheet } = getCategoryData();
  const updatedIcons: CoMapeoIcon[] = [];
  const iconStats = createIconCacheStats();

  categories.forEach((category, index) => {
    const categoryName = category[0];
    const backgroundColor = backgroundColors[index]?.[0] || "#6d44d9";
    const presetSlug = createPresetSlug(categoryName, index);
    const existingIcon = iconMap.get(presetSlug);

    const iconSource =
      existingIcon?.svg ||
      generateNewIcon(categoryName, backgroundColor, presetSlug);

    const savedUrl = saveIconToFolderWithCaching(
      iconsFolder,
      categoryName,
      presetSlug,
      iconSource,
      suffixes,
      backgroundColor,
      true,
      zipBlobs,
      undefined,
      iconStats,
    );

    if (shouldUpdateIconCell(savedUrl, shouldWriteToDrive)) {
      updateIconUrlInSheet(categoriesSheet, index + 2, 2, savedUrl);
    } else {
      getScopedLogger("DriveService").info(
        `[ICON] Skipping sheet update for ${presetSlug} (URL length: ${savedUrl ? savedUrl.length : 0})`,
      );
    }

    if (savedUrl && savedUrl.trim() !== "") {
      updatedIcons.push({
        name: presetSlug,
        svg: savedUrl,
      });
    }
  });

  pruneUnusedIconHashEntries();
  logIconCacheStats(iconStats, shouldWriteToDrive);
  if (iconsFolder) {
    deduplicateFolderFilesByName(iconsFolder);
  }
  return updatedIcons;
}

function saveIconToFolderWithCaching(
  folder: MaybeDriveFolder,
  displayName: string,
  presetSlug: string,
  iconSvg: string,
  suffixes: string[],
  backgroundColor: string,
  shouldWriteToDrive: boolean,
  zipBlobs?: GoogleAppsScript.Base.Blob[],
  depth = 0,
  stats?: IconCacheStats,
): string {
  getScopedLogger("DriveService").info(`[ICON] Saving icon with caching for ${displayName} (slug: ${presetSlug})`);

  const { iconContent, mimeType } = getIconContent({
    svg: iconSvg,
    name: presetSlug,
  });

  if (!iconContent) {
    getScopedLogger("DriveService").warn(`[ICON] Failed to resolve icon content for ${displayName}, generating fallback`);
    const defaultBackground = backgroundColor || "#6d44d9";

    if (depth > 0) {
      getScopedLogger("DriveService").error(`[ICON] Unable to generate icon for ${displayName} after retry, using fallback icon`);
      if (stats) {
        stats.generatedFallbacks++;
      }
      return FALLBACK_ICON_SVG;
    }

    const fallbackSvg = generateNewIcon(
      displayName,
      defaultBackground,
      presetSlug,
    );
    return saveIconToFolderWithCaching(
      folder,
      displayName,
      presetSlug,
      fallbackSvg,
      suffixes,
      defaultBackground,
      shouldWriteToDrive,
      zipBlobs,
      depth + 1,
      stats,
    );
  }

  const effectiveSuffixes = suffixes.length > 0 ? suffixes : [""];
  const allowSkipDriveWrite = Boolean(zipBlobs);
  let resolvedUrl: string | null = null;

  for (const suffix of effectiveSuffixes) {
    const variantUrl = saveIconVariantWithCaching(
      folder,
      presetSlug,
      suffix,
      iconContent,
      mimeType,
      zipBlobs,
      allowSkipDriveWrite,
      true,
      stats,
    );

    if (!resolvedUrl && variantUrl) {
      resolvedUrl = variantUrl;
    }
  }

  if (!resolvedUrl) {
    resolvedUrl = deriveInMemoryIconUrl(iconSvg, iconContent, mimeType, presetSlug, stats);
  }

  if (!resolvedUrl) {
    getScopedLogger("DriveService").warn(`[ICON] Failed to determine icon URL for ${displayName}, returning fallback icon`);
    if (stats) {
      stats.generatedFallbacks++;
    }
    return FALLBACK_ICON_SVG;
  }

  getScopedLogger("DriveService").info(`[ICON] Icon ready for ${displayName}: ${resolvedUrl}`);
  return resolvedUrl;
}

function saveIconVariantWithCaching(
  folder: MaybeDriveFolder,
  presetSlug: string,
  suffix: string,
  iconContent: string,
  mimeType: string,
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  allowSkipDriveWrite: boolean,
  shouldWriteToDrive: boolean,
  stats?: IconCacheStats,
): string | null {
  const sanitizedSize = suffix ? suffix.replace("-", "") : "";
  const propertyKey = getIconHashPropertyKey(presetSlug, sanitizedSize);
  markIconHashUsage(propertyKey);
  const iconHash = computeIconContentHash(iconContent, mimeType);
  const existingMetadata = getIconHashMetadata(propertyKey);

  if (!shouldWriteToDrive) {
    pushIconContentToZip(zipBlobs, iconContent, mimeType, presetSlug, sanitizedSize);
    if (
      existingMetadata &&
      existingMetadata.hash === iconHash &&
      existingMetadata.fileUrl
    ) {
      persistIconHashMetadata(propertyKey, {
        ...existingMetadata,
        updatedAt: new Date().toISOString(),
      });
      getScopedLogger("DriveService").info(
        `[ICON] Reused cached icon metadata for ${presetSlug}${suffix ? suffix : ""} (in-memory mode)`,
      );
      if (stats) {
        stats.reusedFromCache++;
      }
      return existingMetadata.fileUrl;
    }

    persistIconHashMetadata(propertyKey, {
      hash: iconHash,
      updatedAt: new Date().toISOString(),
    });
    return null;
  }

  if (
    allowSkipDriveWrite &&
    existingMetadata &&
    existingMetadata.hash === iconHash
  ) {
    const reusedUrl = reuseExistingIconVariant(
      propertyKey,
      existingMetadata,
      presetSlug,
      sanitizedSize,
      mimeType,
      iconContent,
      zipBlobs,
      folder,
      iconHash,
    );
    if (reusedUrl) {
      if (stats) {
        stats.reusedFromCache++;
      }
      getScopedLogger("DriveService").info(
        `[ICON] Reused cached icon for ${presetSlug}${suffix ? suffix : ""}`,
      );
      return reusedUrl;
    }
    getScopedLogger("DriveService").warn(
      `[ICON] Cached icon metadata invalid for ${presetSlug}${suffix ? suffix : ""}, regenerating`,
    );
  }

  if (!folder) {
    throw new Error("Icons folder is undefined while Drive writes are enabled");
  }

  const file = createIconFile(
    folder,
    presetSlug,
    sanitizedSize,
    iconContent,
    mimeType,
    zipBlobs,
  );
  const fileUrl = file.getUrl();
  if (stats) {
    stats.driveWrites++;
  }
  persistIconHashMetadata(propertyKey, {
    hash: iconHash,
    fileId: file.getId(),
    fileUrl,
    updatedAt: new Date().toISOString(),
  });
  return fileUrl;
}

function deriveInMemoryIconUrl(
  originalIconValue: string,
  iconContent: string,
  mimeType: string,
  presetSlug: string,
  stats?: IconCacheStats,
): string | null {
  if (originalIconValue && originalIconValue.startsWith("data:image")) {
    if (stats) {
      stats.inMemoryUrls++;
    }
    return originalIconValue;
  }

  if (mimeType === MimeType.SVG) {
    if (stats) {
      stats.inMemoryUrls++;
    }
    return `data:image/svg+xml,${encodeURIComponent(iconContent)}`;
  }

  if (mimeType === MimeType.PNG) {
    const base64 = Utilities.base64Encode(
      Utilities.newBlob(iconContent, mimeType, `${presetSlug}.png`).getBytes(),
    );
    if (stats) {
      stats.inMemoryUrls++;
    }
    return `data:image/png;base64,${base64}`;
  }

  return null;
}

function shouldUpdateIconCell(savedUrl: string, shouldWriteToDrive: boolean): boolean {
  if (!savedUrl) {
    return false;
  }

  if (shouldWriteToDrive) {
    return true;
  }

  if (savedUrl.startsWith("https://")) {
    return true;
  }

  return savedUrl.length <= 50000;
}

function reuseExistingIconVariant(
  propertyKey: string,
  metadata: IconHashMetadata,
  presetSlug: string,
  sanitizedSize: string,
  mimeType: string,
  iconContent: string,
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  destinationFolder: MaybeDriveFolder,
  iconHash: string,
): string | null {
  if (!metadata.fileId) {
    getScopedLogger("DriveService").warn(
      `[ICON] Cached icon metadata missing file ID for ${presetSlug}${sanitizedSize ? '-' + sanitizedSize : ''}, regenerating`,
    );
    return null;
  }

  try {
    const existingFile = DriveApp.getFileById(metadata.fileId);
    applyDefaultSharing(existingFile);
    ensureIconFileLinkedToFolder(
      destinationFolder,
      existingFile,
      presetSlug,
      sanitizedSize,
      mimeType,
    );
    const fileUrl = existingFile.getUrl();

    pushIconContentToZip(zipBlobs, iconContent, mimeType, presetSlug, sanitizedSize);

    persistIconHashMetadata(propertyKey, {
      ...metadata,
      hash: iconHash,
      fileId: existingFile.getId(),
      fileUrl,
      updatedAt: new Date().toISOString(),
    });

    return fileUrl;
  } catch (error) {
    getScopedLogger("DriveService").warn(
      `[ICON] Cached icon file ${metadata.fileId} inaccessible: ${error.message}`,
    );
    return null;
  }
}

function ensureIconFileLinkedToFolder(
  folder: MaybeDriveFolder,
  file: GoogleAppsScript.Drive.File,
  presetSlug: string,
  sanitizedSize: string,
  mimeType: string,
): void {
  if (!folder) {
    return;
  }

  try {
    const expectedName = buildIconFileName(presetSlug, sanitizedSize, mimeType);
    const matches = folder.getFilesByName(expectedName);
    while (matches.hasNext()) {
      const existing = matches.next();
      if (existing.getId() === file.getId()) {
        return;
      }
    }

    folder.addFile(file);

    if (file.getName() !== expectedName) {
      getScopedLogger("DriveService").warn(
        `[ICON] Reused icon file name mismatch for ${presetSlug}${sanitizedSize ? '-' + sanitizedSize : ''}: expected ${expectedName}, found ${file.getName()}`,
      );
    }
  } catch (error) {
    getScopedLogger("DriveService").warn(
      `[ICON] Failed to attach cached icon ${file.getId()} to raw build folder: ${error.message}`,
    );
  }
}

function pushIconContentToZip(
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  iconContent: string,
  mimeType: string,
  presetSlug: string,
  sanitizedSize: string,
): void {
  if (!zipBlobs) {
    return;
  }

  const fileName = buildIconFileName(presetSlug, sanitizedSize, mimeType);
  const blob = Utilities.newBlob(iconContent, mimeType, fileName);
  const targetName = `icons/${fileName}`;
  for (let i = zipBlobs.length - 1; i >= 0; i--) {
    if (zipBlobs[i].getName && zipBlobs[i].getName() === targetName) {
      zipBlobs.splice(i, 1);
    }
  }
  zipBlobs.push(blob.copyBlob().setName(targetName));
}

function getDefaultSharingPolicy(): SharingPolicy {
  if (cachedSharingPolicy) {
    return cachedSharingPolicy;
  }

  try {
    const rawPolicy = PropertiesService.getScriptProperties().getProperty(
      SHARING_POLICY_PROPERTY_KEY,
    );
    if (!rawPolicy) {
      cachedSharingPolicy = DEFAULT_SHARING_POLICY;
      return cachedSharingPolicy;
    }

    const normalized = rawPolicy.toLowerCase();
    if (normalized === "anyone" || normalized === "inherit" || normalized === "off") {
      cachedSharingPolicy = normalized as SharingPolicy;
      return cachedSharingPolicy;
    }

    getScopedLogger("DriveService").warn(
      `[SHARING] Unrecognized sharing policy "${rawPolicy}", falling back to "${DEFAULT_SHARING_POLICY}"`,
    );
    cachedSharingPolicy = DEFAULT_SHARING_POLICY;
  } catch (error) {
    getScopedLogger("DriveService").warn(`[SHARING] Failed to read sharing policy: ${error.message}`);
    cachedSharingPolicy = DEFAULT_SHARING_POLICY;
  }

  return cachedSharingPolicy;
}

function applyDefaultSharing(
  target: GoogleAppsScript.Drive.Folder | GoogleAppsScript.Drive.File | null,
): void {
  if (!target) {
    return;
  }

  const policy = getDefaultSharingPolicy();
  if (policy === "off") {
    return;
  }

  try {
    if (policy === "anyone") {
      target.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return;
    }

    if (policy === "inherit") {
      applySpreadsheetSharingToTarget(target);
    }
  } catch (error) {
    getScopedLogger("DriveService").warn(
      `[SHARING] Failed to apply ${policy} sharing to ${target.getName ? target.getName() : "target"}: ${error.message}`,
    );
  }
}

function applySpreadsheetSharingToTarget(
  target: GoogleAppsScript.Drive.Folder | GoogleAppsScript.Drive.File,
): void {
  try {
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const spreadsheetFile = DriveApp.getFileById(spreadsheetId);
    const access = spreadsheetFile.getSharingAccess();
    const permission = spreadsheetFile.getSharingPermission();
    target.setSharing(access, permission);
  } catch (error) {
    getScopedLogger("DriveService").warn(`[SHARING] Unable to inherit spreadsheet sharing: ${error.message}`);
  }
}

function computeIconContentHash(
  iconContent: string,
  mimeType: string,
): string {
  const blob = Utilities.newBlob(iconContent, mimeType);
  return computeBlobMd5Digest(blob);
}

function buildIconFileName(
  presetSlug: string,
  sanitizedSize: string,
  mimeType: string,
): string {
  const extension = mimeType === MimeType.SVG ? "svg" : "png";
  const suffixPart = sanitizedSize ? `-${sanitizedSize}` : "";
  const baseSlug = sanitizeIconSlug(presetSlug) || presetSlug;
  return `${baseSlug}${suffixPart}.${extension}`;
}

function createIconCacheStats(): IconCacheStats {
  return {
    reusedFromCache: 0,
    driveWrites: 0,
    inMemoryUrls: 0,
    generatedFallbacks: 0,
  };
}

function logIconCacheStats(
  stats: IconCacheStats,
  shouldWriteToDrive: boolean,
): void {
  getScopedLogger("DriveService").info(
    `[ICON] Cache summary (${shouldWriteToDrive ? "Drive" : "In-memory"} mode): reused=${stats.reusedFromCache}, driveWrites=${stats.driveWrites}, inlineUrls=${stats.inMemoryUrls}, fallbacks=${stats.generatedFallbacks}`,
  );
}

function getIconHashPropertyKey(
  presetSlug: string,
  sanitizedSize: string,
): string {
  const sizeKey = sanitizedSize || "default";
  return `${ICON_HASH_PROPERTY_PREFIX}${presetSlug}::${sizeKey}`;
}

function getOrCreateSubFolder(
  parent: GoogleAppsScript.Drive.Folder,
  name: string,
  allowExisting = true,
): GoogleAppsScript.Drive.Folder {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    if (!allowExisting) {
      throw new Error(`Folder "${name}" already exists under ${parent.getName()}`);
    }
    return folders.next();
  }
  return parent.createFolder(name);
}

function deduplicateFolderFilesByName(folder: GoogleAppsScript.Drive.Folder): void {
  const seen: Record<string, GoogleAppsScript.Drive.File> = {};
  const duplicates: GoogleAppsScript.Drive.File[] = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    if (seen[name]) {
      duplicates.push(file);
    } else {
      seen[name] = file;
    }
  }

  duplicates.forEach((file) => {
    try {
      getScopedLogger("DriveService").info(`[ICON] Removing duplicate file from Drive: ${file.getName()} (${file.getId()})`);
      folder.removeFile(file);
      file.setTrashed(true);
    } catch (error) {
      getScopedLogger("DriveService").warn(`[ICON] Failed to remove duplicate file ${file.getId()}: ${error.message}`);
    }
  });
}

function markIconHashUsage(propertyKey: string): void {
  iconHashKeysUsedThisRun[propertyKey] = true;
}

function resetIconHashTracking(): void {
  iconHashCache = null;
  iconHashKeysUsedThisRun = {};
}

function getIconHashCache(): Record<string, IconHashMetadata> {
  if (iconHashCache) {
    return iconHashCache;
  }

  const cache: Record<string, IconHashMetadata> = {};
  const properties = PropertiesService.getScriptProperties().getProperties();

  Object.keys(properties)
    .filter((key) => key.startsWith(ICON_HASH_PROPERTY_PREFIX))
    .forEach((key) => {
      try {
        cache[key] = JSON.parse(properties[key]) as IconHashMetadata;
      } catch (error) {
        getScopedLogger("DriveService").warn(`[ICON] Failed to parse icon hash metadata for key "${key}": ${error.message}`);
      }
    });

  iconHashCache = cache;
  return iconHashCache;
}

function getIconHashMetadata(propertyKey: string): IconHashMetadata | null {
  const cache = getIconHashCache();
  return cache[propertyKey] || null;
}

function persistIconHashMetadata(
  propertyKey: string,
  metadata: IconHashMetadata,
): void {
  const cache = getIconHashCache();
  cache[propertyKey] = metadata;
  PropertiesService.getScriptProperties().setProperty(
    propertyKey,
    JSON.stringify(metadata),
  );
  markIconHashUsage(propertyKey);
}

function pruneUnusedIconHashEntries(): void {
  const cache = getIconHashCache();
  const propertiesService = PropertiesService.getScriptProperties();

  Object.keys(cache).forEach((key) => {
    if (!iconHashKeysUsedThisRun[key]) {
      propertiesService.deleteProperty(key);
      delete cache[key];
    }
  });
}

function clearIconHashCache(): void {
  const propertiesService = PropertiesService.getScriptProperties();
  const keys = Object.keys(propertiesService.getProperties()).filter((key) =>
    key.startsWith(ICON_HASH_PROPERTY_PREFIX),
  );

  keys.forEach((key) => {
    propertiesService.deleteProperty(key);
  });

  resetIconHashTracking();
  getScopedLogger("DriveService").info("[ICON] Cleared icon hash cache");
}

function savePresets(
  presets: CoMapeoPreset[],
  presetsFolder: MaybeDriveFolder,
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  shouldWriteToDrive: boolean,
): void {
  for (const preset of presets) {
    const presetJson = JSON.stringify(preset, null, 2);
    const fileName = `${preset.icon}.json`;
    const blob = Utilities.newBlob(presetJson, MimeType.PLAIN_TEXT, fileName);
    if (shouldWriteToDrive && presetsFolder) {
      const file = presetsFolder.createFile(blob);
      applyDefaultSharing(file);
    }
    if (zipBlobs) {
      zipBlobs.push(blob.copyBlob().setName(`presets/${fileName}`));
    }
  }
}

function saveFields(
  fields: CoMapeoField[],
  fieldsFolder: MaybeDriveFolder,
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  shouldWriteToDrive: boolean,
): void {
  for (const field of fields) {
    const fieldJson = JSON.stringify(field, null, 2);
    const fileName = `${field.tagKey}.json`;
    const blob = Utilities.newBlob(fieldJson, MimeType.PLAIN_TEXT, fileName);
    if (shouldWriteToDrive && fieldsFolder) {
      const file = fieldsFolder.createFile(blob);
      applyDefaultSharing(file);
    }
    if (zipBlobs) {
      zipBlobs.push(blob.copyBlob().setName(`fields/${fileName}`));
    }
  }
}

function saveMessages(
  messages: CoMapeoTranslations,
  messagesFolder: MaybeDriveFolder,
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  shouldWriteToDrive: boolean,
): void {
  for (const [lang, langMessages] of Object.entries(messages)) {
    const messagesJson = JSON.stringify(langMessages, null, 2);
    const fileName = `${lang}.json`;
    const blob = Utilities.newBlob(messagesJson, MimeType.PLAIN_TEXT, fileName);
    if (shouldWriteToDrive && messagesFolder) {
      const file = messagesFolder.createFile(blob);
      applyDefaultSharing(file);
    }
    if (zipBlobs) {
      zipBlobs.push(blob.copyBlob().setName(`messages/${fileName}`));
    }
  }
}

function saveMetadataAndPackage(
  config: CoMapeoConfig,
  rootFolder: MaybeDriveFolder,
  zipBlobs: GoogleAppsScript.Base.Blob[] | undefined,
  shouldWriteToDrive: boolean,
): void {
  const metadataBlob = Utilities.newBlob(
    JSON.stringify(config.metadata, null, 2),
    MimeType.PLAIN_TEXT,
    "metadata.json",
  );
  if (shouldWriteToDrive && rootFolder) {
    const metadataFile = rootFolder.createFile(metadataBlob);
    applyDefaultSharing(metadataFile);
  }
  if (zipBlobs) {
    zipBlobs.push(metadataBlob.copyBlob().setName("metadata.json"));
  }

  const packageBlob = Utilities.newBlob(
    JSON.stringify(config.packageJson, null, 2),
    MimeType.PLAIN_TEXT,
    "package.json",
  );
  if (shouldWriteToDrive && rootFolder) {
    const packageFile = rootFolder.createFile(packageBlob);
    applyDefaultSharing(packageFile);
  }
  if (zipBlobs) {
    zipBlobs.push(packageBlob.copyBlob().setName("package.json"));
  }
}

/**
 * Cleans up (deletes) a Drive folder and all its contents.
 * Used for error recovery when config generation fails.
 *
 * @param folderId - The ID of the folder to delete
 */
function cleanupDriveFolder(folderId: string | null): void {
  if (!folderId) {
    getScopedLogger("DriveService").info("[CLEANUP] No folder ID provided, skipping cleanup");
    return;
  }

  try {
    getScopedLogger("DriveService").info("[CLEANUP] Attempting to delete folder with ID:", folderId);
    const folder = DriveApp.getFolderById(folderId);
    const folderName = folder.getName();

    // Move to trash instead of permanent delete
    folder.setTrashed(true);

    getScopedLogger("DriveService").info(`[CLEANUP] ✅ Successfully trashed folder: ${folderName} (ID: ${folderId})`);
  } catch (error) {
    // Log error but don't throw - cleanup failure shouldn't block error handling
    getScopedLogger("DriveService").error("[CLEANUP] ⚠️  Failed to cleanup folder:", error.message);
  }
}
