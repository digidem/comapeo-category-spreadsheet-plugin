/// <reference path="shared.ts" />

/**
 * Validates HTML content for common issues that could cause "Malformed HTML content" errors
 * @param html - The HTML string to validate
 * @returns Object with isValid flag and errors array
 */
function validateHtmlContent(html: string): {
  isValid: boolean;
  errors: string[];
} {
  const errors: string[] = [];

  // Check for basic HTML structure issues
  if (!html || typeof html !== "string") {
    errors.push("HTML content is empty or not a string");
    return { isValid: false, errors };
  }

  // Check for unclosed tags (basic validation)
  const tagStack: string[] = [];
  const selfClosingTags = new Set([
    "img",
    "br",
    "hr",
    "input",
    "meta",
    "link",
    "area",
    "base",
    "col",
    "embed",
    "param",
    "source",
    "track",
    "wbr",
  ]);

  // Match opening and closing tags
  const tagRegex = /<\/?([a-zA-Z][a-zA-Z0-9]*)\b[^>]*>/g;
  let match;

  while ((match = tagRegex.exec(html)) !== null) {
    const fullTag = match[0];
    const tagName = match[1].toLowerCase();

    // Skip self-closing tags
    if (selfClosingTags.has(tagName) || fullTag.endsWith("/>")) {
      continue;
    }

    // Check if it's a closing tag
    if (fullTag.startsWith("</")) {
      if (tagStack.length === 0) {
        errors.push(
          `Closing tag </${tagName}> found without matching opening tag`,
        );
      } else {
        const lastTag = tagStack.pop();
        if (lastTag !== tagName) {
          errors.push(
            `Mismatched tags: Expected </${lastTag}>, found </${tagName}>`,
          );
        }
      }
    } else {
      // Opening tag
      tagStack.push(tagName);
    }
  }

  // Check for unclosed tags
  if (tagStack.length > 0) {
    errors.push(
      `Unclosed tags: ${tagStack.map((tag) => `<${tag}>`).join(", ")}`,
    );
  }

  // Check for common HTML errors
  if (html.includes("<script>") && !html.includes("</script>")) {
    errors.push("Unclosed <script> tag detected");
  }

  if (html.includes("<style>") && !html.includes("</style>")) {
    errors.push("Unclosed <style> tag detected");
  }

  // Check for unescaped special characters in attribute values
  const attrRegex = /(\w+)="([^"]*)"/g;
  let attrMatch;
  while ((attrMatch = attrRegex.exec(html)) !== null) {
    const attrValue = attrMatch[2];
    if (attrValue.includes("<") && !attrValue.startsWith("data:")) {
      errors.push(
        `Unescaped '<' in attribute ${attrMatch[1]}="${attrValue}" - should use &lt;`,
      );
    }
  }

  // Check for malformed attribute syntax (e.g., style="value";> instead of style="value">)
  const malformedAttrRegex = /(\w+)="([^"]*)";>/g;
  let malformedMatch;
  while ((malformedMatch = malformedAttrRegex.exec(html)) !== null) {
    errors.push(
      `Malformed attribute syntax: ${malformedMatch[0]} - semicolon should be inside quotes or removed`,
    );
  }

  // NOTE: Unclosed quote validation removed due to false positives
  // The regex pattern was incorrectly capturing attributes without their closing quotes,
  // causing all valid attributes to be flagged as errors.
  // Other validations (tag matching, script/style tags, etc.) are sufficient.

  // Check for multiple DOCTYPE declarations
  const doctypeCount = (html.match(/<!DOCTYPE/gi) || []).length;
  if (doctypeCount > 1) {
    errors.push("Multiple DOCTYPE declarations found");
  }

  return {
    isValid: errors.length === 0,
    errors,
  };
}

/**
 * Validates HTML before showing a dialog to prevent "Malformed HTML content" errors
 * @param html - The HTML string to validate
 * @param context - Context description for error messages (e.g., "Language Selection Dialog")
 * @throws Error if HTML is malformed
 */
function validateDialogHtml(html: string, context: string = "Dialog"): void {
  const validation = validateHtmlContent(html);

  if (!validation.isValid) {
    const errorMessage = `HTML validation failed for ${context}:\n${validation.errors.join("\n")}`;
    console.error(errorMessage);
    throw new Error(
      `Malformed HTML detected in ${context}. Please check the console for details.`,
    );
  }

  // Additional checks specific to Google Apps Script dialogs
  if (html.length > 500000) {
    console.warn(
      `HTML content for ${context} is very large (${html.length} characters). This may cause performance issues.`,
    );
  }
}

/**
 * Test function to validate HTML dialog generation
 * Run this from the Apps Script editor to test HTML validation
 */
function testHtmlValidation(): void {
  console.log("=== Testing HTML Validation ===");

  // Test cases
  const testCases = [
    {
      name: "Valid HTML",
      html: "<p>Hello <strong>world</strong></p>",
      shouldPass: true,
    },
    {
      name: "Unclosed tag",
      html: "<p>Hello <strong>world</p>",
      shouldPass: false,
    },
    {
      name: "Mismatched tags",
      html: "<div><p>Content</div></p>",
      shouldPass: false,
    },
    {
      name: "Unclosed script tag",
      html: "<script>console.log('test')",
      shouldPass: false,
    },
    {
      name: "Valid self-closing tags",
      html: "<p>Line 1<br/>Line 2<img src='test.png' /></p>",
      shouldPass: true,
    },
    {
      name: "Unescaped < in attribute",
      html: '<p data-value="test<value">Content</p>',
      shouldPass: false,
    },
    {
      name: "Valid data URI",
      html: '<img src="data:image/svg+xml,%3Csvg%3E" />',
      shouldPass: true,
    },
    {
      name: "Multiple DOCTYPE declarations",
      html: "<!DOCTYPE html><!DOCTYPE html><html></html>",
      shouldPass: false,
    },
    {
      name: "Malformed attribute with semicolon",
      html: '<ol style="text-align: left";><li>Item</li></ol>',
      shouldPass: false,
    },
    // NOTE: Unclosed quote validation removed due to false positives
    // The test case for unclosed quotes has been removed
    {
      name: "Valid complex HTML",
      html: "<!DOCTYPE html><html><head><style>body { color: red; }</style></head><body><p>Test</p></body></html>",
      shouldPass: true,
    },
  ];

  let passed = 0;
  let failed = 0;

  testCases.forEach((testCase) => {
    console.log(`\nTesting: ${testCase.name}`);
    const validation = validateHtmlContent(testCase.html);

    if (validation.isValid === testCase.shouldPass) {
      console.log(`✅ PASS: ${testCase.name}`);
      passed++;
    } else {
      console.log(`❌ FAIL: ${testCase.name}`);
      console.log(`  Expected: ${testCase.shouldPass ? "valid" : "invalid"}`);
      console.log(`  Got: ${validation.isValid ? "valid" : "invalid"}`);
      if (validation.errors.length > 0) {
        console.log(`  Errors: ${validation.errors.join(", ")}`);
      }
      failed++;
    }
  });

  console.log("\n=== Test Results ===");
  console.log(`Passed: ${passed}/${testCases.length}`);
  console.log(`Failed: ${failed}/${testCases.length}`);

  if (failed === 0) {
    console.log("✅ All tests passed!");
  } else {
    console.log(`❌ ${failed} test(s) failed`);
  }
}

/**
 * INTEGRATION TEST: Test actual dialog HTML generation
 * This validates the real HTML that would be shown to users
 */
function testDialogHtmlGeneration(): void {
  console.log("\n=== Testing Dialog HTML Generation ===");

  let totalTests = 0;
  let passedTests = 0;

  // Test 1: Simple dialog
  try {
    totalTests++;
    const html = generateDialog("Test Title", "<p>Test message</p>");
    validateDialogHtml(html, "Test Dialog");
    console.log("✅ Simple dialog HTML is valid");
    passedTests++;
  } catch (error) {
    console.log("❌ Simple dialog HTML is INVALID:", error);
  }

  // Test 2: Dialog with button
  try {
    totalTests++;
    const html = generateDialog(
      "Test",
      "<p>Message</p>",
      "Click",
      "https://example.com",
    );
    validateDialogHtml(html, "Dialog with Button");
    console.log("✅ Dialog with button HTML is valid");
    passedTests++;
  } catch (error) {
    console.log("❌ Dialog with button HTML is INVALID:", error);
  }

  // Test 3: Dialog with function button
  try {
    totalTests++;
    const html = generateDialog(
      "Test",
      "<p>Message</p>",
      "Submit",
      null,
      "submitForm",
    );
    validateDialogHtml(html, "Dialog with Function");
    console.log("✅ Dialog with function button HTML is valid");
    passedTests++;
  } catch (error) {
    console.log("❌ Dialog with function button HTML is INVALID:", error);
  }

  // Test 4: Dialog with special characters (should be escaped)
  try {
    totalTests++;
    const title = 'Test <> & "Title"';
    const message = "<p>" + escapeHtml('Message with <> & "quotes"') + "</p>";
    const html = generateDialog(title, message);
    validateDialogHtml(html, "Dialog with Special Chars");
    console.log("✅ Dialog with special characters HTML is valid");
    passedTests++;
  } catch (error) {
    console.log("❌ Dialog with special characters HTML is INVALID:", error);
  }

  console.log(`\n=== Dialog Generation Test Results ===`);
  console.log(`Passed: ${passedTests}/${totalTests}`);

  if (passedTests === totalTests) {
    console.log("✅ All dialog generation tests passed!");
  } else {
    console.log(
      `❌ ${totalTests - passedTests} dialog generation test(s) failed`,
    );
  }
}

/**
 * Run all HTML validation tests
 */
function runAllHtmlValidationTests(): void {
  testHtmlValidation();
  testDialogHtmlGeneration();
  console.log("\n=== All HTML Validation Tests Complete ===");
}
