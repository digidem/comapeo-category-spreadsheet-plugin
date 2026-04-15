/**
 * Lint parity regression tests
 *
 * Run these from the Apps Script editor to verify lint and builder parsing stay aligned.
 * Usage:
 *   runLintParityTests();
 *
 * Depends on globals from src/builders/payloadBuilder.ts:
 * normalizeFieldTokens(), splitTranslatedOptions() (loaded earlier by GAS alphabetical order).
 */

function testFieldTokenParity(): boolean {
  console.log("=== testFieldTokenParity ===");

  const cases: Array<{ input: any; expected: string[] }> = [
    { input: "River, Forest, Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River;Forest;Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River\nForest\nMountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River•Forest•Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River·Forest·Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "River，Forest，Mountain", expected: ["River", "Forest", "Mountain"] },
    { input: "  River  ,  Forest  ", expected: ["River", "Forest"] },
    { input: "", expected: [] },
    { input: undefined, expected: [] },
    { input: null, expected: [] },
    { input: ["River", "Forest"], expected: ["River", "Forest"] },
    { input: "Single", expected: ["Single"] },
  ];

  for (const { input, expected } of cases) {
    const result = normalizeFieldTokens(input);
    const matches =
      result.length === expected.length &&
      result.every((value, index) => value === expected[index]);
    if (!matches) {
      console.error(
        `FAIL: normalizeFieldTokens(${JSON.stringify(input)}) => ` +
          `[${result.map((value) => `"${value}"`).join(", ")}] expected ` +
          `[${expected.map((value) => `"${value}"`).join(", ")}]`,
      );
      return false;
    }
  }

  console.log(`PASS: ${cases.length}/${cases.length} field token parity cases passed`);
  return true;
}

function testCanonicalOptionParity(): boolean {
  console.log("=== testCanonicalOptionParity ===");

  const cases: Array<{
    input: string;
    expected: Array<{ value: string; label: string }>;
  }> = [
    {
      input: "Oak, Pine, Maple",
      expected: [
        { value: "oak", label: "Oak" },
        { value: "pine", label: "Pine" },
        { value: "maple", label: "Maple" },
      ],
    },
    {
      input: "oak:Oak Tree, pine:Pine Tree",
      expected: [
        { value: "oak", label: "Oak Tree" },
        { value: "pine", label: "Pine Tree" },
      ],
    },
    {
      input: "  A  ,  B  ",
      expected: [
        { value: "a", label: "A" },
        { value: "b", label: "B" },
      ],
    },
    {
      input: "single",
      expected: [{ value: "single", label: "single" }],
    },
    {
      input: "",
      expected: [],
    },
    {
      input: "val:Label",
      expected: [{ value: "val", label: "Label" }],
    },
    {
      input: "key:Label:Extra",
      expected: [{ value: "key", label: "Label:Extra" }],
    },
  ];

  for (const { input, expected } of cases) {
    const result = parseCanonicalOptions(input);
    const matches =
      result.length === expected.length &&
      result.every(
        (entry, index) =>
          entry.value === expected[index].value &&
          entry.label === expected[index].label,
      );
    if (!matches) {
      console.error(
        `FAIL: parseCanonicalOptions("${input}") => ` +
          `[${result
            .map((entry) => `{value:"${entry.value}",label:"${entry.label}"}`)
            .join(", ")}] expected ` +
          `[${expected
            .map((entry) => `{value:"${entry.value}",label:"${entry.label}"}`)
            .join(", ")}]`,
      );
      return false;
    }
  }

  console.log(`PASS: ${cases.length}/${cases.length} canonical option parity cases passed`);
  return true;
}

function testTranslationDelimiterParity(): boolean {
  console.log("=== testTranslationDelimiterParity ===");

  const cases: Array<{ input: string; expectedCount: number }> = [
    { input: "School, Hospital, Farm", expectedCount: 3 },
    { input: "School;Hospital;Farm", expectedCount: 3 },
    { input: "学校、病院、農場", expectedCount: 3 },
    { input: "学校，医院，农场", expectedCount: 3 },
    { input: "A,B;C，D、E", expectedCount: 5 },
    { input: "  A  ,  B  ", expectedCount: 2 },
    { input: "", expectedCount: 0 },
    { input: "Single", expectedCount: 1 },
  ];

  for (const { input, expectedCount } of cases) {
    const result = splitTranslatedOptions(input);
    if (result.length !== expectedCount) {
      console.error(
        `FAIL: splitTranslatedOptions("${input}") => ${result.length} parts ` +
          `[${result.map((value) => `"${value}"`).join(", ")}] expected ${expectedCount}`,
      );
      return false;
    }
  }

  console.log(`PASS: ${cases.length}/${cases.length} translation delimiter parity cases passed`);
  return true;
}

function runLintParityTests(): void {
  console.log("=== Lint Parity Regression Tests ===");

  const tests = [
    { name: "Field Token Parity", fn: testFieldTokenParity },
    { name: "Canonical Option Parity", fn: testCanonicalOptionParity },
    { name: "Translation Delimiter Parity", fn: testTranslationDelimiterParity },
  ];

  let passed = 0;
  let failed = 0;

  for (const test of tests) {
    try {
      if (test.fn()) {
        passed++;
      } else {
        failed++;
      }
    } catch (error) {
      failed++;
      console.error(`FAIL: ${test.name} threw an exception - ${error}`);
    }
  }

  console.log(`runLintParityTests: ${passed}/${tests.length} passed`);

  if (failed > 0) {
    throw new Error(`${failed} lint parity test(s) failed`);
  }
}
