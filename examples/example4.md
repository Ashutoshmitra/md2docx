# Code Style Guide

## Purpose

This style guide provides coding conventions for our development team. Consistent coding practices improve code readability, reduce bugs, and make maintenance easier.

## General Principles

* Write code for humans first, computers second
* Favor readability over cleverness
* Be consistent
* Follow the principle of least surprise
* Write self-documenting code

## Naming Conventions

### Variables

* Use descriptive names that reveal intent
* Use camelCase for variables and functions
* Use PascalCase for classes
* Use UPPER_SNAKE_CASE for constants

**Good examples:**

```javascript
const userAge = 25;
const isValid = true;
const MAX_RETRY_COUNT = 3;
```

**Bad examples:**

```javascript
const a = 25;
const valid = true;
const Retries = 3;
```

### Functions

* Use verbs for function names that describe the action
* Keep functions small and focused on a single task
* Aim for 20 lines or less per function

**Good example:**

```javascript
function calculateTotalPrice(items, taxRate) {
  const subtotal = sumItemPrices(items);
  const tax = calculateTax(subtotal, taxRate);
  return subtotal + tax;
}
```

## Code Formatting

### Indentation

* Use 2 spaces for indentation, not tabs
* Maintain consistent indentation throughout the project

### Line Length

* Limit lines to 80-100 characters
* Break long lines sensibly at logical points

### Spacing

* Add spaces around operators
* Add a space after commas
* No space between function name and opening parenthesis
* One space after keywords like if, for, while

**Example:**

```javascript
if (condition) {
  doSomething();
} else {
  doSomethingElse();
}

for (let i = 0; i < items.length; i++) {
  processItem(items[i]);
}
```

## Comments

* Write comments to explain why, not what
* Keep comments updated when code changes
* Use JSDoc style for function documentation

**Example:**

```javascript
/**
 * Calculates the total price including tax
 * @param {Array} items - Array of item objects with price property
 * @param {number} taxRate - Tax rate as decimal (e.g., 0.07 for 7%)
 * @returns {number} Total price including tax
 */
function calculateTotalPrice(items, taxRate) {
  // Implementation...
}
```

## Error Handling

* Never swallow errors silently
* Always provide meaningful error messages
* Handle edge cases explicitly

**Example:**

```javascript
try {
  const data = JSON.parse(jsonString);
  processData(data);
} catch (error) {
  logger.error('Failed to parse JSON data', { error, jsonString });
  throw new Error('Invalid data format');
}
```

## Testing

* Write tests before or alongside code
* Test both success and failure cases
* Mock external dependencies

**Example:**

```javascript
describe('calculateTotalPrice', () => {
  it('should calculate total with tax', () => {
    const items = [{ price: 10 }, { price: 20 }];
    const result = calculateTotalPrice(items, 0.1);
    expect(result).toBe(33); // 30 + 3 tax
  });
  
  it('should handle empty items array', () => {
    expect(calculateTotalPrice([], 0.1)).toBe(0);
  });
});
```

## File Organization

* One logical component per file
* Group related files in directories
* Keep file sizes manageable (under 400 lines)
* Follow consistent import ordering

## Version Control

* Write clear, descriptive commit messages
* Make small, focused commits
* Pull and rebase before pushing changes
* Use feature branches for new development