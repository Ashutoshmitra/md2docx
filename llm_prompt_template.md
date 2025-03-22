# Markdown Formatting Guidelines for CBS Word Template Conversion

## Instructions for LLM

You are tasked with generating Markdown content that will be converted into a Microsoft Word document using a specific converter. Your output must follow these strict formatting guidelines to ensure proper styling in the final Word document.

## Heading Structure

Use the proper heading hierarchy with # symbols:

```markdown
# Document Title (H1)
## Major Section (H2)
### Sub-section (H3)
#### Minor sub-section (H4)
```

- Do not use underlines (like === or ---) for headings
- Do not add manual numbering to headings (the Word template will handle this)
- Add a blank line before and after each heading

## Lists

For unordered lists, use hyphens (-) consistently:

```markdown
- First item
- Second item
  - Nested item
  - Another nested item
- Third item
```

For ordered lists, use numbers followed by periods:

```markdown
1. First item
2. Second item
   1. Nested numbered item
   2. Another nested numbered item
3. Third item
```

- Do not use asterisks (*) or plus signs (+) for bullet points
- Indent nested list items with exactly 2 spaces
- Do not add extra blank lines between list items
- Add a blank line before the first item and after the last item

## Text Formatting

- **Bold text**: Use double asterisks (`**bold**`)
- *Italic text*: Use single asterisks (`*italic*`)
- `Code`: Use backticks for inline code (`` `code` ``)
- For emphasis, prefer bold over italic when appropriate

## Code Blocks

Use triple backticks with the language specified:

```markdown
​```python
def example_function():
    return "Hello, world!"
​```
```

- Add a blank line before and after code blocks
- Specify the language when applicable

## Tables

For tables, use the standard pipe syntax:

```markdown
| Header 1 | Header 2 | Header 3 |
|----------|----------|----------|
| Cell 1   | Cell 2   | Cell 3   |
| Cell 4   | Cell 5   | Cell 6   |
```

- Align the pipes in your markdown for better readability
- Add a blank line before and after tables

## Line Breaks and Paragraphs

- Use a blank line to create a new paragraph
- Do not use trailing spaces or backslashes for line breaks
- Do not use HTML `<br>` tags
- Avoid soft returns within paragraphs

## Other Guidelines

- Do not use HTML formatting tags (the converter handles only Markdown)
- Do not use custom CSS or styling
- Keep lines reasonably short (under 100 characters)
- Use consistent spacing throughout the document
- Avoid indenting paragraphs with spaces

## Example Document

```markdown
# Sample Document Title

## Introduction

This is an introduction paragraph. It should be clear and concise, providing an overview of the document.

## Key Concepts

This section outlines the key concepts discussed in this document.

### First Concept

- Point one about the first concept
- Point two about the first concept
  - Additional detail
  - More detail

### Second Concept

1. First step in the process
2. Second step in the process
   1. Sub-step A
   2. Sub-step B
3. Third step in the process

## Technical Specifications

The following table shows the technical specifications:

| Parameter | Value | Unit |
|-----------|-------|------|
| Speed     | 150   | km/h |
| Weight    | 50    | kg   |
| Height    | 2.5   | m    |

## Code Example

Here is a simple code example:

​```python
def calculate_total(items):
    total = 0
    for item in items:
        total += item['price']
    return total
​```

## Conclusion

This concludes the document. Thank you for reading.
```

Please follow these guidelines carefully to ensure that your generated Markdown can be properly converted to a Word document using the CBS template system.