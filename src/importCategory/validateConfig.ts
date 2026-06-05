/** Maximum number of validation errors to report. Prevents memory issues on corrupt archives. */
const MAX_VALIDATION_ERRORS = 20;

/** Valid document types per CoMapeo spec */
const VALID_DOCUMENT_TYPES = ["observation", "track"];

/** Valid field types per CoMapeo spec */
const VALID_FIELD_TYPES = ["text", "number", "selectOne", "selectMultiple"];

/**
 * Validates the structure of imported configuration data before applying to spreadsheet.
 * Catches malformed .comapeocat files early with clear error messages.
 *
 * Checks conform to the canonical CoMapeo category schema defined in
 * `package/src/schema/category.js` and `package/src/schema/field.js`.
 *
 * @param config - The parsed configuration object to validate
 * @returns Object with isValid flag and array of error messages
 */
function validateImportedConfig(config: unknown): { isValid: boolean; errors: string[] } {
  const errors: string[] = [];

  // Must be an object
  if (!config || typeof config !== "object") {
    return { isValid: false, errors: ["Configuration data is not a valid object"] };
  }

  const cfg = config as Record<string, unknown>;

  // Check required top-level keys
  if (!Array.isArray(cfg.presets)) {
    errors.push("Missing or invalid 'presets' array");
  }

  if (!Array.isArray(cfg.fields)) {
    errors.push("Missing or invalid 'fields' array");
  }

  if (!Array.isArray(cfg.icons)) {
    errors.push("Missing or invalid 'icons' array");
  }

  if (!cfg.messages || typeof cfg.messages !== "object" || Array.isArray(cfg.messages)) {
    errors.push("Missing or invalid 'messages' object");
  }

  // Validate presets structure (if present)
  if (Array.isArray(cfg.presets)) {
    for (let i = 0; i < cfg.presets.length && errors.length < MAX_VALIDATION_ERRORS; i++) {
      const preset = cfg.presets[i];
      if (!preset || typeof preset !== "object") {
        errors.push(`Preset at index ${i} is not a valid object`);
        continue;
      }
      const p = preset as Record<string, unknown>;

      // name is required
      if (typeof p.name !== "string" || !p.name.trim()) {
        errors.push(`Preset '${p.id || `index ${i}`}' missing required 'name' property`);
      }

      // appliesTo is required per spec
      if (!Array.isArray(p.appliesTo) || p.appliesTo.length === 0) {
        errors.push(`Preset '${p.name || `index ${i}`}' missing required 'appliesTo' array (expected ['observation'] and/or ['track'])`);
      } else {
        const invalidTypes = (p.appliesTo as string[]).filter(
          (t) => !VALID_DOCUMENT_TYPES.includes(t),
        );
        if (invalidTypes.length > 0 && errors.length < MAX_VALIDATION_ERRORS) {
          errors.push(`Preset '${p.name || `index ${i}`}' has unrecognized appliesTo values: ${invalidTypes.join(", ")}`);
        }
      }

      // tags is required per spec (must be non-empty object)
      if (!p.tags || typeof p.tags !== "object" || Array.isArray(p.tags) || Object.keys(p.tags as Record<string, unknown>).length === 0) {
        if (errors.length < MAX_VALIDATION_ERRORS) {
          errors.push(`Preset '${p.name || `index ${i}`}' missing required 'tags' object (must have at least one entry)`);
        }
      }
    }
  }

  // Validate fields structure (if present)
  if (Array.isArray(cfg.fields)) {
    for (let i = 0; i < cfg.fields.length && errors.length < MAX_VALIDATION_ERRORS; i++) {
      const field = cfg.fields[i];
      if (!field || typeof field !== "object") {
        errors.push(`Field at index ${i} is not a valid object`);
        continue;
      }
      const f = field as Record<string, unknown>;

      // label is required
      if (typeof f.label !== "string" || !f.label.trim()) {
        errors.push(`Field '${f.id || `index ${i}`}' missing required 'label' property`);
      }

      // tagKey is required per spec
      if (typeof f.tagKey !== "string" || !f.tagKey.trim()) {
        if (errors.length < MAX_VALIDATION_ERRORS) {
          errors.push(`Field '${f.id || `index ${i}`}' missing required 'tagKey' property`);
        }
      }

      // type should be a valid enum value
      if (f.type !== undefined && !VALID_FIELD_TYPES.includes(f.type as typeof VALID_FIELD_TYPES[number])) {
        if (errors.length < MAX_VALIDATION_ERRORS) {
          errors.push(`Field '${f.id || `index ${i}`}' has unrecognized type '${f.type}' (expected one of: ${VALID_FIELD_TYPES.join(", ")})`);
        }
      }
    }
  }

  // Validate metadata (optional but should be object if present)
  if (cfg.metadata !== undefined && cfg.metadata !== null) {
    if (typeof cfg.metadata !== "object" || Array.isArray(cfg.metadata)) {
      if (errors.length < MAX_VALIDATION_ERRORS) {
        errors.push("'metadata' should be an object if present");
      }
    }
  }

  // Check for completely empty config (only when all arrays are valid but empty)
  const allArraysValid = Array.isArray(cfg.presets) && Array.isArray(cfg.fields) && Array.isArray(cfg.icons);
  const messagesObj = cfg.messages && typeof cfg.messages === "object" && !Array.isArray(cfg.messages)
    ? (cfg.messages as Record<string, unknown>)
    : {};
  const hasMessages = Object.keys(messagesObj).length > 0;
  if (allArraysValid && (cfg.presets as unknown[]).length === 0 && (cfg.fields as unknown[]).length === 0 && (cfg.icons as unknown[]).length === 0 && !hasMessages) {
    if (errors.length < MAX_VALIDATION_ERRORS) {
      errors.push("Configuration appears empty — no presets, fields, icons, or messages found");
    }
  }

  // Add summary if errors were capped — push the summary so no real error is silently dropped
  if (errors.length >= MAX_VALIDATION_ERRORS) {
    errors.push(`... and more issues may exist (showing first ${MAX_VALIDATION_ERRORS})`);
  }

  return {
    isValid: errors.length === 0,
    errors,
  };
}
