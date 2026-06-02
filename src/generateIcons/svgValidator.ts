/**
 * SVG validation layer for icon processing
 *
 * Provides validation functions to check SVG format, Drive access,
 * and content validity before processing icons.
 */

/**
 * Result of an SVG validation operation
 */
interface SvgValidationResult {
  /** Whether validation passed */
  valid: boolean;

  /** Error message if validation failed */
  error?: string;

  /** Additional context about the validation */
  context?: Record<string, any>;
}

/**
 * Validates that a string contains valid SVG format
 *
 * Checks for:
 * - Presence of <svg> opening tag
 * - Basic XML structure
 * - Valid namespace declaration
 *
 * @param svgContent - The SVG string to validate
 * @returns Validation result with details
 */
function validateSvgFormat(svgContent: string): SvgValidationResult {
  if (!svgContent || svgContent.trim() === "") {
    return {
      valid: false,
      error: "SVG content is empty",
    };
  }

  // Check if it's a data URI
  if (svgContent.startsWith("data:image/svg+xml")) {
    try {
      // Extract and decode the SVG content from data URI
      const svgPart = svgContent.replace(/^data:image\/svg\+xml[^,]*,/, "");
      const decoded = decodeURIComponent(svgPart);
      return validateSvgFormat(decoded); // Recursively validate the decoded content
    } catch (error) {
      return {
        valid: false,
        error: "Failed to decode SVG data URI",
        context: { originalError: error.message },
      };
    }
  }

  // Check if it's a Drive URL (these are validated separately)
  if (svgContent.startsWith("https://drive.google.com")) {
    return {
      valid: true,
      context: { type: "drive-url" },
    };
  }

  // Basic SVG tag check
  if (!svgContent.includes("<svg")) {
    return {
      valid: false,
      error: "Missing <svg> opening tag",
    };
  }

  // Check for closing tag
  if (!svgContent.includes("</svg>")) {
    return {
      valid: false,
      error: "Missing </svg> closing tag",
    };
  }

  // Try to parse as XML to check structure
  try {
    const xml = XmlService.parse(svgContent);
    const root = xml.getRootElement();

    // Verify root element is 'svg'
    if (root.getName() !== "svg") {
      return {
        valid: false,
        error: `Root element is '${root.getName()}', expected 'svg'`,
      };
    }

    // Check for namespace (recommended but not strictly required)
    const namespace = root.getNamespace();
    if (!namespace || !namespace.getURI()) {
      console.warn(
        "SVG missing namespace declaration - may cause rendering issues",
      );
      // Don't fail, just warn
    }

    return {
      valid: true,
      context: {
        hasNamespace: Boolean(namespace && namespace.getURI()),
        rootElement: root.getName(),
      },
    };
  } catch (error) {
    return {
      valid: false,
      error: `Invalid XML structure: ${error.message}`,
      context: { parseError: error.message },
    };
  }
}

/**
 * Validates that a Google Drive file is accessible
 *
 * @param fileId - The Google Drive file ID
 * @returns Validation result with file details
 */
function validateDriveAccess(fileId: string): SvgValidationResult {
  if (!fileId || fileId.trim() === "") {
    return {
      valid: false,
      error: "File ID is empty",
    };
  }

  try {
    const file = DriveApp.getFileById(fileId);

    // Try to access basic properties to ensure we have read access
    const name = file.getName();
    const mimeType = file.getMimeType();

    return {
      valid: true,
      context: {
        fileName: name,
        mimeType: mimeType,
        fileId: fileId,
      },
    };
  } catch (error) {
    // Parse the error to provide more specific feedback
    const errorMsg = error.message || String(error);

    if (errorMsg.includes("not found") || errorMsg.includes("does not exist")) {
      return {
        valid: false,
        error: "File not found",
        context: { fileId, originalError: errorMsg },
      };
    }

    if (errorMsg.includes("permission") || errorMsg.includes("access")) {
      return {
        valid: false,
        error: "Permission denied - you don't have access to this file",
        context: { fileId, originalError: errorMsg },
      };
    }

    return {
      valid: false,
      error: `Drive access failed: ${errorMsg}`,
      context: { fileId, originalError: errorMsg },
    };
  }
}

/**
 * Validates an icon URL format and accessibility
 *
 * Supports:
 * - Inline SVG markup (<svg>...</svg>)
 * - Google Drive URLs (https://drive.google.com/file/d/...)
 * - Data URIs (data:image/svg+xml,...)
 * - External HTTPS URLs
 * - Plain text search terms (e.g., "river", "building", "tree")
 *
 * @param iconUrl - The URL, data URI, or search term to validate
 * @returns Validation result with URL details
 */
function validateIconUrl(iconUrl: string): SvgValidationResult {
  if (!iconUrl || iconUrl.trim() === "") {
    return {
      valid: false,
      error: "Icon URL is empty",
    };
  }

  const trimmed = iconUrl.trim();

  // Data URI validation
  if (trimmed.startsWith("data:image/svg+xml")) {
    // Extract and validate the SVG content
    return validateSvgFormat(trimmed);
  }

  // Drive URL validation
  if (trimmed.startsWith("https://drive.google.com/file/d/")) {
    try {
      const fileId = trimmed.split("/d/")[1].split("/")[0];
      return validateDriveAccess(fileId);
    } catch (error) {
      return {
        valid: false,
        error: "Invalid Drive URL format",
        context: { url: trimmed, originalError: error.message },
      };
    }
  }

  // External HTTPS URL - allow and handle during processing
  if (
    trimmed.startsWith("https://") &&
    !trimmed.startsWith("https://drive.google.com")
  ) {
    return {
      valid: true,
      context: {
        type: "external-url",
        url: trimmed,
      },
    };
  }

  // Inline SVG markup validation
  if (trimmed.includes("<svg") && trimmed.includes("</svg>")) {
    return validateSvgFormat(trimmed);
  }

  // Plain text search term - any non-empty string is valid
  // These will be passed to the icon API as search terms
  return {
    valid: true,
    context: {
      type: "search-term",
      searchTerm: trimmed,
    },
  };
}

/**
 * Validates SVG content for proper structure and namespace
 *
 * This is a more thorough validation than validateSvgFormat.
 * It checks:
 * - Valid XML structure
 * - Proper SVG namespace
 * - Presence of content (paths, shapes, etc.)
 * - Valid viewBox attribute
 *
 * @param svgContent - The SVG content to validate
 * @returns Validation result with content analysis
 */
function validateSvgContent(svgContent: string): SvgValidationResult {
  // First run basic format validation
  const formatValidation = validateSvgFormat(svgContent);
  if (!formatValidation.valid) {
    return formatValidation;
  }

  // Skip detailed validation for Drive URLs (validated separately)
  if (svgContent.startsWith("https://drive.google.com")) {
    return { valid: true, context: { type: "drive-url" } };
  }

  // Decode data URI if needed
  let cleanSvg = svgContent;
  if (svgContent.startsWith("data:image/svg+xml")) {
    try {
      const svgPart = svgContent.replace(/^data:image\/svg\+xml[^,]*,/, "");
      cleanSvg = decodeURIComponent(svgPart);
    } catch (error) {
      return {
        valid: false,
        error: "Failed to decode SVG data URI",
        context: { originalError: error.message },
      };
    }
  }

  try {
    const xml = XmlService.parse(cleanSvg);
    const root = xml.getRootElement();

    // Check namespace
    const namespace = root.getNamespace();
    const hasNamespace = Boolean(namespace && namespace.getURI());

    if (!hasNamespace) {
      return {
        valid: false,
        error:
          'SVG missing namespace declaration (xmlns="http://www.w3.org/2000/svg")',
        context: { recommendation: "Add xmlns attribute to <svg> tag" },
      };
    }

    // Check for content (at least one child element)
    const children = root.getChildren();
    if (children.length === 0) {
      return {
        valid: false,
        error: "SVG has no content (no child elements)",
        context: { recommendation: "Add paths, shapes, or other SVG elements" },
      };
    }

    // Check for viewBox (recommended)
    const viewBox = root.getAttribute("viewBox");
    if (!viewBox) {
      console.warn("SVG missing viewBox attribute - may affect scaling");
      // Don't fail, just note it
    }

    return {
      valid: true,
      context: {
        hasNamespace: true,
        hasViewBox: Boolean(viewBox),
        childCount: children.length,
        childTypes: children.map((c) => c.getName()),
      },
    };
  } catch (error) {
    return {
      valid: false,
      error: `SVG content validation failed: ${error.message}`,
      context: { parseError: error.message },
    };
  }
}

/**
 * Validates an icon from a spreadsheet cell
 *
 * Handles different cell types:
 * - CellImage objects
 * - String URLs (Drive links)
 * - SVG data URIs
 *
 * @param cellValue - The value from the spreadsheet cell
 * @returns Validation result with cell type info
 */
function validateCellIcon(cellValue: any): SvgValidationResult {
  if (!cellValue) {
    return {
      valid: false,
      error: "Cell is empty",
    };
  }

  // Check if it's a CellImage object
  if (typeof cellValue === "object" && cellValue.toString() === "CellImage") {
    try {
      const url = cellValue.getUrl();
      if (!url) {
        return {
          valid: false,
          error: "CellImage has no URL",
        };
      }

      return {
        valid: true,
        context: {
          type: "cell-image",
          url: url,
        },
      };
    } catch (error) {
      return {
        valid: false,
        error: `Failed to access CellImage: ${error.message}`,
        context: { originalError: error.message },
      };
    }
  }

  // Check if it's a string (URL or data URI)
  if (typeof cellValue === "string") {
    return validateIconUrl(cellValue);
  }

  return {
    valid: false,
    error: `Unsupported cell value type: ${typeof cellValue}`,
    context: { valueType: typeof cellValue },
  };
}

/**
 * Quick validation check for batch processing
 *
 * Performs lightweight checks to quickly identify obvious issues
 * without full XML parsing.
 *
 * @param iconValue - The icon value to check
 * @returns True if the icon appears valid for processing
 */
function quickValidateIcon(iconValue: any): boolean {
  if (!iconValue) return false;

  // CellImage objects are valid
  if (typeof iconValue === "object" && iconValue.toString() === "CellImage") {
    return true;
  }

  if (typeof iconValue === "string") {
    const trimmed = iconValue.trim();

    // Empty strings are invalid
    if (trimmed === "") return false;

    // Drive URLs are valid
    if (trimmed.startsWith("https://drive.google.com/file/d/")) {
      return true;
    }

    // External HTTPS URLs are valid
    if (trimmed.startsWith("https://")) {
      return true;
    }

    // Data URIs are valid
    if (trimmed.startsWith("data:image/svg+xml")) {
      return true;
    }

    // Check for basic SVG structure
    if (trimmed.includes("<svg") && trimmed.includes("</svg>")) {
      return true;
    }

    // Unknown format
    return false;
  }

  return false;
}
