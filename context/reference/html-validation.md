# HTML Validation Strategy

## Overview

This document describes the HTML validation strategy implemented to prevent "Malformed HTML content" errors in Google Apps Script dialogs.

## Problem

The error "Malformed HTML content" occurs when Google Apps Script attempts to render a dialog with malformed HTML. This can happen when:

1. **Unclosed HTML tags** - Tags that are opened but never closed
2. **Mismatched tags** - Opening and closing tags don't match (e.g., `<div>...</p>`)
3. **Invalid tag nesting** - Tags nested in improper order
4. **Unescaped special characters** - Characters like `<`, `>` in attribute values
5. **Multiple DOCTYPE declarations** - More than one DOCTYPE in a document
6. **Script/style tag issues** - Unclosed `<script>` or `<style>` tags

**Note**: Unclosed quote validation in attributes was removed due to regex false positives. Use proper escaping and client-side DOM rendering to prevent quote issues.

## Solution Architecture

### 1. HTML Validation Function (`src/lint/htmlValidation.ts`)

```typescript
validateHtmlContent(html: string): { isValid: boolean; errors: string[] }
```

**Core validation checks:**
- Tag matching using a stack-based algorithm
- Self-closing tag detection (img, br, hr, input, etc.)
- Script and style tag validation
- Unescaped special characters in attribute values
- Malformed attribute syntax (e.g., `style="value";>`)
- DOCTYPE declaration count

**Not validated** (due to false positives):
- Unclosed quotes in attributes - Use client-side DOM rendering instead

**Returns:**
- `isValid`: Boolean indicating if HTML is valid
- `errors`: Array of error messages describing issues

### 2. Dialog Validation Function (`src/lint/htmlValidation.ts`)

```typescript
validateDialogHtml(html: string, context?: string): void
```

**Purpose:** Pre-flight validation before showing dialogs
**Throws:** Error if HTML is malformed
**Warnings:** Large HTML content (>500KB)

### 3. Integration in Dialog Generation (`src/dialog.ts`)

```typescript
generateDialog(...) {
  // ... generate HTML ...

  // Validate before returning
  if (typeof validateDialogHtml === "function") {
    try {
      validateDialogHtml(html, title || "Dialog");
    } catch (error) {
      // Return safe fallback dialog
      return safeFallbackHtml;
    }
  }

  return html;
}
```

### 4. Message Sanitization Helper (`src/dialog.ts`)

```typescript
validateAndSanitizeMessage(messageLines: string[]): string
```

**Purpose:** Validate and sanitize message content before passing to `generateDialog`
**Behavior:**
- Validates existing HTML tags in message lines
- Escapes invalid HTML as plain text
- Wraps plain text in `<p>` tags
- Returns properly structured HTML

## Usage Guidelines

### For Dialog Creation

```typescript
// BAD - No validation
const html = generateDialog(title, `<p>Message ${userInput}`);
SpreadsheetApp.getUi().showModalDialog(...);

// GOOD - Validation built-in
const html = generateDialog(title, message); // Validates automatically
SpreadsheetApp.getUi().showModalDialog(...);
```

### For Message Construction

```typescript
// BAD - Manual HTML construction
const message = messages.map(msg => `<p>${msg}</p>`).join('\n');

// GOOD - Use sanitization helper
const message = validateAndSanitizeMessage(messages);
```

### For User Input

```typescript
// BAD - Unescaped user input
const message = `<p>Hello ${userName}</p>`;

// GOOD - Escape user input
const message = `<p>Hello ${escapeHtml(userName)}</p>`;
```

## Validation Algorithm

### Tag Stack Algorithm

```
1. Initialize empty stack
2. For each HTML tag match:
   a. If self-closing (br, img, etc.) → skip
   b. If closing tag → pop from stack and verify match
   c. If opening tag → push to stack
3. After all matches, stack should be empty
4. If not empty → unclosed tags error
```

### Example

```html
<div>
  <p>Hello <strong>world</strong></p>
</div>
```

**Stack evolution:**
1. `<div>` → Stack: [div]
2. `<p>` → Stack: [div, p]
3. `<strong>` → Stack: [div, p, strong]
4. `</strong>` → Pop strong, Stack: [div, p]
5. `</p>` → Pop p, Stack: [div]
6. `</div>` → Pop div, Stack: []
7. ✅ Valid (empty stack)

## Testing

### Manual Testing

Run the test function from Apps Script editor:

```typescript
testHtmlValidation();
```

### Test Cases

The `testHtmlValidation()` function includes 10 test cases:

1. ✅ Valid HTML with nested tags
2. ❌ Unclosed tag
3. ❌ Mismatched tags
4. ❌ Unclosed script tag
5. ✅ Valid self-closing tags
6. ❌ Unescaped `<` in attribute
7. ✅ Valid data URI
8. ❌ Multiple DOCTYPE declarations
9. ❌ Malformed attribute with semicolon
10. ✅ Valid complex HTML

**Note**: Unclosed quote test case was removed after validation was disabled due to false positives.

### Expected Output

```
=== Testing HTML Validation ===

Testing: Valid HTML
✅ PASS: Valid HTML

Testing: Unclosed tag
✅ PASS: Unclosed tag

...

=== Test Results ===
Passed: 8/8
✅ All tests passed!
```

## Common Issues and Solutions

### Issue 1: Unclosed Tags

**Problem:**
```typescript
const message = `<p>Hello <strong>world`;
```

**Solution:**
```typescript
const message = `<p>Hello <strong>world</strong></p>`;
```

### Issue 2: User Input in HTML

**Problem:**
```typescript
const message = `<p>${userInput}</p>`;
```

**Solution:**
```typescript
const message = `<p>${escapeHtml(userInput)}</p>`;
```

### Issue 3: Dynamic HTML Generation

**Problem:**
```typescript
let html = '<ul>';
items.forEach(item => {
  html += `<li>${item.name}`; // Missing </li>
});
html += '</ul>';
```

**Solution:**
```typescript
const html = '<ul>' +
  items.map(item => `<li>${escapeHtml(item.name)}</li>`).join('') +
  '</ul>';
```

### Issue 4: Template String Issues

**Problem:**
```typescript
const message = `
  <div>
    <p>Line 1
    <p>Line 2</p>
  </div>
`; // Missing </p> for Line 1
```

**Solution:**
```typescript
const message = `
  <div>
    <p>Line 1</p>
    <p>Line 2</p>
  </div>
`;
```

## Performance Considerations

### Validation Cost

- **Time Complexity:** O(n) where n is HTML string length
- **Space Complexity:** O(d) where d is max nesting depth
- **Typical Performance:** <10ms for dialogs (<50KB)

### When to Skip Validation

In production, validation can be skipped for:
1. Static HTML templates (pre-validated)
2. Performance-critical paths (with monitoring)
3. Trusted internal content (with testing)

**Note:** Current implementation validates all dialogs to prevent runtime errors.

## Debugging

### Enable Verbose Logging

```typescript
console.log("HTML to validate:", html);
const validation = validateHtmlContent(html);
console.log("Validation result:", validation);
```

### Common Error Messages

1. **"Unclosed tags: \<tag\>"** - Tags opened but not closed
2. **"Mismatched tags: Expected \</tag1\>, found \</tag2\>"** - Wrong closing tag
3. **"Closing tag \</tag\> found without matching opening tag"** - Extra closing tag
4. **"Unclosed \<script\> tag detected"** - Script tag not closed
5. **"Unescaped '\<' in attribute"** - Need to use `&lt;`
6. **"Multiple DOCTYPE declarations found"** - Remove duplicate DOCTYPEs

## Integration Checklist

When adding new dialog functions:

- [ ] Use `generateDialog()` for HTML generation
- [ ] Pass sanitized messages using `validateAndSanitizeMessage()`
- [ ] Escape all user input with `escapeHtml()`
- [ ] Test with `testHtmlValidation()` after changes
- [ ] Check console logs for validation warnings
- [ ] Verify dialog displays correctly in UI

## Future Enhancements

### Potential Improvements

1. **HTML Sanitization Library** - Use a more robust HTML parser
2. **Attribute Validation** - Validate attribute names and values
3. **CSS Validation** - Check inline styles for malformed CSS
4. **JavaScript Validation** - Validate inline JavaScript code
5. **Performance Optimization** - Cache validation results for static content
6. **Custom Rules** - Allow projects to define custom validation rules

### Known Limitations

1. **Unclosed Quotes in Attributes** - Not validated due to regex false positives. Use client-side DOM rendering or proper escaping.
2. **Complex HTML** - May not catch all edge cases in deeply nested HTML
3. **Invalid Entities** - Doesn't validate HTML entities (e.g., `&nbsp;`)
4. **CSS in Styles** - Doesn't validate CSS syntax within `<style>` tags
5. **JavaScript in Scripts** - Doesn't validate JavaScript syntax within `<script>` tags

## References

- [Google Apps Script HTML Service](https://developers.google.com/apps-script/guides/html)
- [HTML5 Specification](https://html.spec.whatwg.org/)
- [Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)

## Change Log

- **2025-10-13** - Initial implementation with tag validation, test suite, and documentation
