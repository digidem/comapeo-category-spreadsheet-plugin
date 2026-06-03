/**
 * Validates the structure of imported configuration data before applying to spreadsheet.
 * Catches malformed .comapeocat files early with clear error messages.
 *
 * @param config - The parsed configuration object to validate
 * @returns Object with valid flag and array of error messages
 */
function validateImportedConfig(config: unknown): { valid: boolean; errors: string[] } {
  const errors: string[] = [];

  // Must be an object
  if (!config || typeof config !== "object") {
    return { valid: false, errors: ["Configuration data is not a valid object"] };
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
    for (let i = 0; i < cfg.presets.length; i++) {
      const preset = cfg.presets[i];
      if (!preset || typeof preset !== "object") {
        errors.push(`Preset at index ${i} is not a valid object`);
        continue;
      }
      const p = preset as Record<string, unknown>;
      if (typeof p.name !== "string" || !p.name.trim()) {
        errors.push(`Preset at index ${i} missing 'name' property`);
      }
    }
  }

  // Validate fields structure (if present)
  if (Array.isArray(cfg.fields)) {
    for (let i = 0; i < cfg.fields.length; i++) {
      const field = cfg.fields[i];
      if (!field || typeof field !== "object") {
        errors.push(`Field at index ${i} is not a valid object`);
        continue;
      }
      const f = field as Record<string, unknown>;
      if (typeof f.label !== "string" || !f.label.trim()) {
        errors.push(`Field at index ${i} missing 'label' property`);
      }
    }
  }

  // Validate metadata (optional but should be object if present)
  if (cfg.metadata !== undefined && cfg.metadata !== null) {
    if (typeof cfg.metadata !== "object" || Array.isArray(cfg.metadata)) {
      errors.push("'metadata' should be an object if present");
    }
  }

  // Check for completely empty config (all required arrays empty)
  const presetsArr = Array.isArray(cfg.presets) ? cfg.presets : [];
  const fieldsArr = Array.isArray(cfg.fields) ? cfg.fields : [];
  const iconsArr = Array.isArray(cfg.icons) ? cfg.icons : [];
  if (presetsArr.length === 0 && fieldsArr.length === 0 && iconsArr.length === 0) {
    errors.push("Configuration appears empty — no presets, fields, or icons found");
  }

  return {
    valid: errors.length === 0,
    errors,
  };
}
