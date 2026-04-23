/**
 * @param {import("../../src/schema/defaults.js").DefaultsDeprecatedInput} defaults
 * @return {import("../../src/schema/categorySelection.js").CategorySelectionInput}
 */
export function migrateDefaults(defaults) {
	return {
		observation: [...new Set([...(defaults.point || []), ...(defaults.area || [])])],
		track: defaults.line,
	}
}
