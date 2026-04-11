/// <reference path="../utils.ts" />

/**
 * File parsing functions for the import category functionality.
 * This file contains functions related to parsing extracted files.
 *
 * Uses the slugify helpers from utils.ts
 */

const logParseFiles = getScopedLogger("ParseFiles");

/**
 * Restructures flat translation keys into nested object structure.
 * Converts flat format like "presets.water.name" to nested format like presets.presets.water.name
 *
 * @param flatMessages - Messages object with flat key structure
 * @returns Messages object with nested structure
 */
function restructureTranslations(flatMessages: any): any {
	if (!flatMessages || typeof flatMessages !== "object") {
		logParseFiles.warn("Invalid flatMessages object, returning empty structure");
		return {};
	}

	const toScalar = (input: any, fallback: string): string => {
		if (input === undefined || input === null) {
			return fallback;
		}
		if (typeof input === "string") {
			return input;
		}
		if (typeof input === "object") {
			if (typeof (input as { message?: unknown }).message === "string") {
				return (input as { message: string }).message;
			}
			const nestedMessage = (input as { message?: { label?: string } }).message;
			if (nestedMessage && typeof nestedMessage.label === "string") {
				return nestedMessage.label;
			}
			if (typeof (input as { label?: unknown }).label === "string") {
				return (input as { label: string }).label;
			}
		}
		try {
			return String(input);
		} catch (_error) {
			return fallback;
		}
	};

	const normalizeOption = (rawValue: any, optionId: string) => {
		const label = toScalar(rawValue, "");
		let value = optionId;
		if (rawValue && typeof rawValue === "object") {
			const nestedMessage = (rawValue as { message?: { value?: string } }).message;
			if (nestedMessage && typeof nestedMessage.value === "string") {
				value = nestedMessage.value;
			} else if (typeof (rawValue as { value?: string }).value === "string") {
				value = (rawValue as { value: string }).value;
			}
		}
		return { label, value };
	};

	const mergeOptionMap = (
		container: Record<string, any>,
		source: Record<string, any>,
	) => {
		if (!source || typeof source !== "object") {
			return;
		}
		container.options = container.options || {};
		for (const [optionId, optionValue] of Object.entries(source)) {
			container.options[optionId] = normalizeOption(optionValue, optionId);
		}
	};

	const mergeScalarProps = (
		target: Record<string, any>,
		source: Record<string, any>,
	) => {
		if (!source || typeof source !== "object") {
			return;
		}
		for (const [prop, value] of Object.entries(source)) {
			if (prop === "options") {
				mergeOptionMap(target, value as Record<string, any>);
			} else if (value !== undefined) {
				target[prop] = toScalar(value, "");
			}
		}
	};

	const hasTranslationValue = (input: any): boolean => {
		if (typeof input === "string") {
			return input.trim() !== "";
		}
		if (Array.isArray(input)) {
			return input.some((item) => hasTranslationValue(item));
		}
		if (input && typeof input === "object") {
			if (typeof (input as { label?: string }).label === "string" && (input as { label: string }).label.trim() !== "") {
				return true;
			}
			for (const [key, value] of Object.entries(input)) {
				if (key === "value") {
					continue;
				}
				if (hasTranslationValue(value)) {
					return true;
				}
			}
		}
		return false;
	};

	const restructured: any = {};

	for (const lang of Object.keys(flatMessages)) {
		const langMessages = flatMessages[lang];
		if (!langMessages || typeof langMessages !== "object") {
			logParseFiles.warn(`Invalid messages for language: ${lang}`);
			continue;
		}

		const langResult = {
			presets: {
				presets: {} as Record<string, any>,
				fields: {} as Record<string, any>,
			},
		};

		const nestedPresets =
			langMessages?.presets && typeof langMessages.presets === "object"
				? langMessages.presets
				: undefined;
		const nestedFields =
			langMessages?.fields && typeof langMessages.fields === "object"
				? langMessages.fields
				: undefined;

		if (nestedPresets) {
			if (nestedPresets.presets && typeof nestedPresets.presets === "object") {
				for (const [presetId, presetValue] of Object.entries(nestedPresets.presets)) {
					const targetPreset = langResult.presets.presets[presetId] || {};
					mergeScalarProps(targetPreset, presetValue as Record<string, any>);
					langResult.presets.presets[presetId] = targetPreset;
				}
			}
			if (nestedPresets.categories && typeof nestedPresets.categories === "object") {
				for (const [presetId, presetValue] of Object.entries(nestedPresets.categories)) {
					const targetPreset = langResult.presets.presets[presetId] || {};
					mergeScalarProps(targetPreset, presetValue as Record<string, any>);
					langResult.presets.presets[presetId] = targetPreset;
				}
			}
			if (nestedPresets.fields && typeof nestedPresets.fields === "object") {
				for (const [fieldId, fieldValue] of Object.entries(nestedPresets.fields)) {
					const targetField = langResult.presets.fields[fieldId] || {};
					mergeScalarProps(targetField, fieldValue as Record<string, any>);
					langResult.presets.fields[fieldId] = targetField;
				}
			}
			for (const [maybePresetId, presetValue] of Object.entries(nestedPresets)) {
				if (maybePresetId === "presets" || maybePresetId === "fields" || maybePresetId === "categories") {
					continue;
				}
				if (!presetValue || typeof presetValue !== "object") {
					continue;
				}
				const targetPreset = langResult.presets.presets[maybePresetId] || {};
				mergeScalarProps(targetPreset, presetValue as Record<string, any>);
				langResult.presets.presets[maybePresetId] = targetPreset;
			}
		}

		if (nestedFields) {
			for (const [fieldId, fieldValue] of Object.entries(nestedFields)) {
				const targetField = langResult.presets.fields[fieldId] || {};
				mergeScalarProps(targetField, fieldValue as Record<string, any>);
				langResult.presets.fields[fieldId] = targetField;
			}
		}

		for (const [rawKey, rawValue] of Object.entries(langMessages)) {
			if (typeof rawKey !== "string" || !rawKey.includes(".")) {
				continue;
			}
			const parts = rawKey.split(".");
			if (parts.length < 3) {
				continue;
			}
			const messageValue =
				typeof rawValue === "object" && rawValue !== null && "message" in rawValue
					? (rawValue as any).message
					: rawValue;

			if (parts[0] === "presets") {
				const itemId = parts[1];
				const property = parts[2];
				const targetPreset = langResult.presets.presets[itemId] || {};
				if (parts[3] === "options" && parts.length >= 5) {
					mergeOptionMap(targetPreset, {
						[parts[4]]: messageValue,
					});
				} else {
					targetPreset[property] = toScalar(messageValue, "");
				}
				langResult.presets.presets[itemId] = targetPreset;
			} else if (parts[0] === "fields") {
				const itemId = parts[1];
				const property = parts[2];
				const targetField = langResult.presets.fields[itemId] || {};
				if (parts[3] === "options" && parts.length >= 5) {
					mergeOptionMap(targetField, {
						[parts[4]]: messageValue,
					});
				} else {
					targetField[property] = toScalar(messageValue, "");
				}
				langResult.presets.fields[itemId] = targetField;
			}
		}

		const hasPresets = hasTranslationValue(langResult.presets.presets);
		const hasFields = hasTranslationValue(langResult.presets.fields);
		if (hasPresets || hasFields) {
			restructured[lang] = langResult;
		} else {
			logParseFiles.info(`Skipping language ${lang} with no translation values`);
		}
	}

	logParseFiles.info(
		`Restructured translations for ${Object.keys(restructured).length} languages`,
	);
	return restructured;
}

/**
 * Parses the extracted files to get configuration data.
 * @param files - Array of extracted file blobs
 * @param tempFolder - The temporary folder where files are extracted
 * @param onProgress - Optional progress callback function
 * @returns Configuration data object
 */
function parseExtractedFiles(
	files: GoogleAppsScript.Base.Blob[] | undefined,
	tempFolder: GoogleAppsScript.Drive.Folder,
	onProgress?: (update: { percent: number; stage: string; detail?: string }) => void,
): any {
	// Handle undefined or empty files array
	if (!files || !Array.isArray(files)) {
		logParseFiles.info("No files to parse or files is not an array");
		return {
			metadata: { name: "Empty Configuration" },
			presets: [],
			fields: [],
			messages: {},
			icons: [],
		};
	}

	logParseFiles.info(`Parsing ${files.length} extracted files...`);

	// Initialize configuration data
	const configData: any = {
		metadata: null,
		presets: [],
		fields: [],
		messages: {},
		icons: [],
	};

	// Process each file
	files.forEach((file) => {
		const fileName = file.getName();
		logParseFiles.info(`Processing file: ${fileName}`);

		try {
			// Parse metadata.json
			if (fileName === "metadata.json") {
				const content = JSON.parse(file.getDataAsString());
				configData.metadata = content;
				logParseFiles.info("Parsed metadata.json:", configData.metadata);
			}

			// Parse presets.json (contains categories and fields)
			else if (fileName === "presets.json") {
				const content = JSON.parse(file.getDataAsString());

				// Extract presets (categories)
				if (content.presets) {
					// Handle object format (Mapeo style)
					if (
						typeof content.presets === "object" &&
						!Array.isArray(content.presets)
					) {
						for (const presetId in content.presets) {
							const preset = content.presets[presetId];
							configData.presets.push({
								id: presetId,
								name: preset.name || presetId,
								icon: preset.icon || presetId,
								color: preset.color || "#0000FF",
								fields: preset.fields || [],
							});
						}
					}
				}

				// Extract fields
				if (content.fields) {
					// Handle object format (Mapeo style)
					if (
						typeof content.fields === "object" &&
						!Array.isArray(content.fields)
					) {
						for (const fieldId in content.fields) {
							const field = content.fields[fieldId];

							// Convert field type
							let fieldType = "text";
							if (field.type === "select_one") fieldType = "selectOne";
							else if (field.type === "select_multiple")
								fieldType = "selectMultiple";

						// Convert options
						let options = [];
						if (field.options) {
							if (Array.isArray(field.options)) {
								// Handle array of strings or objects
								options = field.options.map((opt: any, optionIndex: number) => {
									if (typeof opt === "string") {
										return {
											label: opt,
											value: createOptionValue(opt, fieldId, optionIndex),
										};
									} else if (typeof opt === "object" && opt !== null) {
										const optionLabel = opt.label || opt.value || opt.name || String(opt);
										const normalizedValue =
											typeof opt.value === "string" && opt.value.trim() !== ""
												? opt.value
											: createOptionValue(optionLabel, fieldId, optionIndex);
										return {
											label: optionLabel,
											value: normalizedValue,
										};
									}
									const fallbackLabel = String(opt);
									return {
										label: fallbackLabel,
										value: createOptionValue(fallbackLabel, fieldId, optionIndex),
									};
								});
							} else if (typeof field.options === "object") {
								options = Object.entries(field.options).map(
									([key, value]) => ({
										label: typeof value === "string" ? value : key,
										value: key,
									}),
								);
							}
						}

					configData.fields.push({
						id: fieldId,
						tagKey: fieldId,
						label: field.label || fieldId,
						type: fieldType,
						helperText: field.helperText || field.placeholder || "",
						placeholder: field.placeholder || field.helperText || "",
						options: options,
					});
						}
					}
				}

				logParseFiles.info(
					`Parsed presets.json: ${configData.presets.length} presets, ${configData.fields.length} fields`,
				);
			}

			// Parse translations.json
			else if (fileName === "translations.json") {
				const content = JSON.parse(file.getDataAsString());
				// Restructure flat translation keys to nested format
				configData.messages = restructureTranslations(content);
				logParseFiles.info(
					`Parsed and restructured translations.json: ${Object.keys(content).length} languages`,
				);
			}

			// Process icons.svg file
			else if (fileName === "icons.svg") {
				logParseFiles.info("Found icons.svg file, will process sprite");
				// Store the file for later processing
				configData.iconsSvgFile = file;
			}
			// Process icons.png sprite file
			else if (fileName === "icons.png") {
				logParseFiles.info("Found PNG sprite file (icons.png)");
				// Store for detection, but cannot extract without external API
				configData.iconsPngFile = file;
			}
			// Process icons.json metadata file
			else if (fileName === "icons.json") {
				logParseFiles.info("Found icon metadata file (icons.json)");
				// Contains coordinates for PNG sprite extraction
				configData.iconsJsonFile = file;
			}
			// Skip individual icon files here - they'll be processed later by processIconSpriteBlob() and extractPngIcons()
			// DO NOT add placeholder entries as they block proper extraction
			else if (
				fileName.endsWith(".svg") ||
				fileName.endsWith(".png") ||
				fileName.includes("/icons/") ||
				fileName.startsWith("icons/")
			) {
				// Individual icon files are handled by extraction phases:
				// Phase 1 (Primary): processIconSpriteBlob() for SVG sprite extraction
				// Phase 2 (Fallback): extractPngIcons() for PNG files in icons/ directory
				// Adding placeholder entries here would block actual extraction
				logParseFiles.info(`Detected individual icon file: ${fileName} (will be processed later)`);
			}
		} catch (error) {
			logParseFiles.warn(`Error processing file ${fileName}:`, error);
		}
	});

	// Try to extract SVG icons first (preferred format for scalability and display)
	logParseFiles.info("=== ICON EXTRACTION PHASE 1: SVG Sprite (Primary) ===");
	if (configData.iconsSvgFile) {
		logParseFiles.info("Processing icons.svg sprite file");
		logParseFiles.info(`Icons before SVG extraction: ${configData.icons.length}`);
		try {
			// Process the SVG sprite and get icon objects with URLs
			// processIconSpriteBlob returns { icons: [...], errors: [...] }
			const spriteResult = processIconSpriteBlob(
				tempFolder,
				configData.iconsSvgFile,
			);
			const extractedIcons = spriteResult.icons || [];
			logParseFiles.info(`✓ Extracted ${extractedIcons.length} icons from SVG sprite`);

			// Log any errors that occurred during extraction
			if (spriteResult.errors && spriteResult.errors.length > 0) {
				logParseFiles.warn(`⚠️ ${spriteResult.errors.length} icon(s) failed to extract:`);
				spriteResult.errors.forEach((err) => {
					logParseFiles.warn(`  - ${err.symbolId}: ${err.error}`);
				});
			}

			// Add the extracted icons to the configData.icons array
			if (extractedIcons.length > 0) {
				// Create a map of existing icon names to avoid duplicates
				const existingIconNames = new Set<string>();
				configData.icons.forEach((icon: { name: string }) => {
					existingIconNames.add(icon.name);
				});

				if (existingIconNames.size > 0) {
					logParseFiles.info(`Existing icon names: ${Array.from(existingIconNames).join(", ")}`);
				}

				// Add SVG icons (they now take precedence)
				let svgAdded = 0;
				let svgSkipped = 0;
				extractedIcons.forEach(
					(icon: { name: string; svg: string; id: string }) => {
						if (!existingIconNames.has(icon.name)) {
							configData.icons.push(icon);
							svgAdded++;
							logParseFiles.info(`  ✓ Added SVG icon: ${icon.name} (URL: ${icon.svg.substring(0, 50)}...)`);
						} else {
							svgSkipped++;
							logParseFiles.info(`  ⊘ Skipped duplicate SVG icon: ${icon.name}`);
						}
					},
				);

				logParseFiles.info(`Added ${svgAdded} new SVG icons${svgSkipped > 0 ? `, skipped ${svgSkipped} duplicates` : ""}`);
				logParseFiles.info(`Total icons after SVG processing: ${configData.icons.length}`);
			}
		} catch (error) {
			logParseFiles.error("Error processing icons.svg:", error);
		}

		// Remove the temporary file reference
		delete configData.iconsSvgFile;
	} else {
		logParseFiles.info("No icons.svg file found, will try PNG files");
	}

	// Fall back to PNG icons from icons/ directory if SVG extraction didn't find any
	logParseFiles.info("=== ICON EXTRACTION PHASE 2: PNG Files (Fallback) ===");
	logParseFiles.info(`Icons before PNG extraction: ${configData.icons.length}`);
	if (hasPngIconsDirectory(tempFolder)) {
		logParseFiles.info("Found icons/ directory, extracting PNG icons");
		try {
			const pngIcons = extractPngIcons(tempFolder, configData.presets, onProgress);
			logParseFiles.info(`✓ Extracted ${pngIcons.length} PNG icons`);

			if (pngIcons.length > 0) {
				// Create a map of existing icon names to avoid duplicates
				const existingIconNames = new Set<string>();
				configData.icons.forEach((icon: { name: string }) => {
					existingIconNames.add(icon.name);
				});

				if (existingIconNames.size > 0) {
					logParseFiles.info(`Existing icon names (SVG takes precedence): ${Array.from(existingIconNames).join(", ")}`);
				}

				// Add PNG icons that don't already exist (SVG icons take precedence)
				let pngAdded = 0;
				let pngSkipped = 0;
				pngIcons.forEach(
					(icon: { name: string; svg: string; id: string }) => {
						if (!existingIconNames.has(icon.name)) {
							configData.icons.push(icon);
							existingIconNames.add(icon.name);
							pngAdded++;
							logParseFiles.info(`  ✓ Added PNG icon: ${icon.name} (URL: ${icon.svg.substring(0, 50)}...)`);
						} else {
							pngSkipped++;
							logParseFiles.info(`  ⊘ Skipped PNG icon (SVG exists): ${icon.name}`);
						}
					},
				);

				logParseFiles.info(`Added ${pngAdded} new PNG icons${pngSkipped > 0 ? `, skipped ${pngSkipped} (SVG priority)` : ""}`);
				logParseFiles.info(`Total icons after PNG extraction: ${configData.icons.length}`);
			}
		} catch (error) {
			logParseFiles.error("Error extracting PNG icons:", error);
		}
	} else {
		logParseFiles.info("No icons/ directory found, PNG extraction skipped");
	}

	// Check if PNG sprite was detected but not extracted
	logParseFiles.info("=== ICON EXTRACTION COMPLETE ===");
	logParseFiles.info(`Final icon count: ${configData.icons.length}`);

	if (configData.icons.length === 0) {
		if (configData.iconsPngFile) {
			logParseFiles.warn("⚠️ PNG SPRITE DETECTED BUT CANNOT BE EXTRACTED");
			logParseFiles.warn("   Google Apps Script cannot parse PNG sprites (icons.png)");
			logParseFiles.warn("   Reason: No image processing libraries available");
			logParseFiles.warn("");
			logParseFiles.warn("   Options:");
			logParseFiles.warn("   1. Use config with individual PNG files in icons/ directory");
			logParseFiles.warn("   2. Use config with SVG sprite (icons.svg)");
			logParseFiles.warn("   3. Contact support about external API for PNG extraction");

			if (configData.iconsJsonFile) {
				logParseFiles.warn("   Note: icons.json metadata detected but unusable without PNG extraction");
			}
		} else {
			logParseFiles.warn("⚠️ NO ICONS FOUND IN CONFIGURATION");
			logParseFiles.warn("   Expected one of:");
			logParseFiles.warn("   - icons/ directory with individual PNG files");
			logParseFiles.warn("   - icons.svg sprite file");
			logParseFiles.warn("   Please check configuration file structure");
		}
	} else {
		logParseFiles.info(`✓ Successfully extracted ${configData.icons.length} icon(s)`);
	}

	// Clean up temporary file references
	delete configData.iconsSvgFile;
	delete configData.iconsPngFile;
	delete configData.iconsJsonFile;

	return configData;
}
