#!/usr/bin/env python3
"""
MD to DOCX Converter Module
Converts Markdown files to Word document format using a specified template
"""

import os
import re
from docx import Document
import markdown
from bs4 import BeautifulSoup
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_docx_from_readme(readme_path, template_path, output_path):
    """
    Convert a Markdown file to a Word document using a template.
    
    Args:
        readme_path (str): Path to the Markdown file
        template_path (str): Path to the Word template
        output_path (str): Path to save the output Word document
    
    Returns:
        bool: True if conversion was successful, False otherwise
    """
    try:
        # Load the template document
        doc = Document(template_path)
        
        # List all styles in the document for debugging
        print("Available styles in template:")
        for style in doc.styles:
            print(f"- {style.name}")
        
        # Preserve headers and footers by removing only body paragraphs
        body = doc._body
        for element in body._element[:]:
            if element.tag.endswith('p'):
                body._element.remove(element)

        # Read the README file and remove soft carriage returns
        with open(readme_path, 'r', encoding='utf-8') as f:
            readme_content = f.read().replace('\r\n', '\n').replace('\r', '\n')
        
        # Split the markdown content into lines for direct parsing
        markdown_lines = readme_content.split('\n')
        
        # Convert Markdown to HTML for non-list elements
        html_content = markdown.markdown(readme_content, extensions=['extra', 'tables', 'fenced_code'])
        soup = BeautifulSoup(html_content, 'html.parser')

        # Helper function to clean manual numbering from headings
        def clean_heading_text(text):
            # Match patterns like "1.", "1.1", "1.1.1" at the start of a heading
            # Updated to be more aggressive in removing all section numbers
            pattern = r'^(\d+(\.\d+)*(\s+|\.|\s+\.)?)\s*'
            if re.match(pattern, text):
                return re.sub(pattern, '', text)
            return text

        # Helper function to add styled paragraph with formatting support
        def add_paragraph(doc, element, style_name):
            paragraph = doc.add_paragraph()
            
            # Try to set the style
            try:
                paragraph.style = style_name
            except KeyError:
                print(f"Warning: Style '{style_name}' not found in template. Using 'Normal'.")
                paragraph.style = 'Normal'
            
            # Get and clean text for headings
            raw_text = element.get_text().strip()
            if style_name.startswith('Heading'):
                text = clean_heading_text(raw_text)
            else:
                text = raw_text
            
            # Process contents with formatting support
            if element.find(['strong', 'em', 'i', 'b', 'code']):
                # Handle paragraph with formatted text
                for content in element.children:
                    if isinstance(content, str):
                        paragraph.add_run(content)
                    elif content.name == 'strong' or content.name == 'b':
                        run = paragraph.add_run(content.get_text())
                        run.bold = True
                    elif content.name == 'em' or content.name == 'i':
                        run = paragraph.add_run(content.get_text())
                        run.italic = True
                    elif content.name == 'code':
                        run = paragraph.add_run(content.get_text())
                        run.font.name = 'Courier New'
                        run.font.size = Pt(10)
            else:
                # Simple paragraph without formatting
                paragraph.add_run(text)
            
            return paragraph

        # Helper function to add code block
        def add_code_block(doc, code_text, language=None):
            container = doc.add_paragraph()
            container.style = 'Normal'
            
            # Add gray background
            shading_elm = OxmlElement('w:shd')
            shading_elm.set(qn('w:fill'), "F0F0F0")
            container._p.get_or_add_pPr().append(shading_elm)
            
            # Add code content line by line with courier font
            lines = code_text.strip().split('\n')
            for i, line in enumerate(lines):
                if line.strip() == '':
                    container.add_run('\n')
                else:
                    run = container.add_run(line)
                    run.font.name = 'Courier New'
                    run.font.size = Pt(10)
                    
                    # Add newline if not the last line
                    if i < len(lines) - 1:
                        container.add_run('\n')
            
            return container

        # Helper function to add table
        def add_table(doc, table_soup):
            rows = table_soup.find_all('tr')
            if not rows:
                return
            
            max_cols = max(len(row.find_all(['th', 'td'])) for row in rows)
            table = doc.add_table(rows=len(rows), cols=max_cols)
            table.style = 'Table Grid'
            
            for i, row in enumerate(rows):
                cells = row.find_all(['th', 'td'])
                for j, cell in enumerate(cells):
                    if j < max_cols:
                        cell_text = cell.get_text().strip()
                        paragraph = table.rows[i].cells[j].paragraphs[0]
                        if cell.name == 'th':
                            run = paragraph.add_run(cell_text)
                            run.bold = True
                        else:
                            paragraph.add_run(cell_text)

        # ======== IDENTIFY LIST BLOCKS FOR LATER PROCESSING ========
        def identify_list_blocks(lines):
            list_blocks = []
            current_block = None
            in_code_block = False
            
            for i, line in enumerate(lines):
                # Check if this line starts or ends a code block
                if line.strip().startswith('```'):
                    in_code_block = not in_code_block
                    if current_block is not None:
                        # This code block is part of the list item
                        current_block.append(i)
                    continue
                    
                # If we're inside a code block, add to current list block if it exists
                if in_code_block and current_block is not None:
                    current_block.append(i)
                    continue
                    
                # Check if line is a list item
                is_list_item = re.match(r'^(\s*)[-*+](.+)$', line) or re.match(r'^(\s*)\d+\.(.+)$', line)
                
                if is_list_item:
                    if current_block is None:
                        # Start a new list block
                        current_block = [i]
                    else:
                        # Continue the existing block
                        current_block.append(i)
                elif current_block is not None:
                    if line.strip() == '':
                        # Empty line - add to current block as it might be part of list formatting
                        current_block.append(i)
                    elif line.strip().startswith('```'):
                        # Beginning of a code block - consider part of the list
                        current_block.append(i)
                    else:
                        # Non-empty, non-list line after a list - end the block if it's not indented
                        # Check if it's indented (part of a list item content)
                        leading_space = len(line) - len(line.lstrip())
                        if leading_space > 0:  # If indented, consider it part of the list
                            current_block.append(i)
                        else:  # Not indented, end the list block
                            list_blocks.append((current_block[0], i-1))
                            current_block = None
            
            # Don't forget the last block if we're still in one
            if current_block is not None:
                list_blocks.append((current_block[0], len(lines)-1))
                    
            return list_blocks

        # Process inline formatting like bold, italic, etc.
        def process_inline_formatting(text):
            # Clean up any problematic markdown syntax first
            # Fix issues with inconsistent asterisks or markdown syntax
            text = re.sub(r'\*\*\*', '**', text)  # Replace *** with **
            text = re.sub(r'\\\*', '*', text)     # Replace \* with *
            text = re.sub(r'\*\*\*', '**', text)  # One more pass to catch any newly formed ***
            
            # Convert the text to HTML to handle the formatting
            html_snippet = markdown.markdown(text)
            
            # Parse the HTML
            soup = BeautifulSoup(html_snippet, 'html.parser')
            
            # If it's a simple paragraph with formatting, return the paragraph's HTML
            if soup.p:
                return soup.p
            
            # Otherwise, return the original text
            return text

        # Function to parse and process a list block
        # When processing list blocks, add code block handling:
        def process_list_block(doc, lines, start_line, end_line):
            # Extract all list items with their indentation levels
            items = []
            current_list_item = None
            in_code_block = False
            code_block_lines = []
            
            i = start_line
            while i <= end_line:
                line = lines[i].rstrip()
                
                # Check if this starts/ends a code block
                if line.strip().startswith('```'):
                    if not in_code_block:
                        # Starting a code block
                        in_code_block = True
                        code_block_start = i
                        language = line.strip()[3:]
                    else:
                        # Ending a code block
                        in_code_block = False
                        
                        # Add the code block to the document
                        code_text = '\n'.join(code_block_lines)
                        add_code_block(doc, code_text, language)
                        code_block_lines = []
                        
                    i += 1
                    continue
                    
                if in_code_block:
                    # Collect code lines
                    code_block_lines.append(line)
                    i += 1
                    continue
                
                # If we're not in a code block, process list items normally
                if not line.strip():
                    i += 1
                    continue
                    
                # Check if it's a list item
                ul_match = re.match(r'^(\s*)[-*+](.+)$', line)
                ol_match = re.match(r'^(\s*)\d+\.(.+)$', line)
                
                if ul_match or ol_match:
                    # If we had a previous list item, add it now
                    if current_list_item is not None:
                        items.append(current_list_item)
                        
                    match = ul_match or ol_match
                    indent = len(match.group(1))
                    content = match.group(2).strip()
                    is_ordered = bool(ol_match)
                    
                    # Process bold formatting in list items
                    content = process_inline_formatting(content)
                    
                    # Store this list item
                    current_list_item = (indent, content, is_ordered, i)
                    
                i += 1
            
            # Add the last list item if any
            if current_list_item is not None:
                items.append(current_list_item)
                    
            # Process items with correct nesting
            if items:
                create_nested_list_items(doc, items)
                return True
            return False

        # Create list items with correct nesting based on indentation
        def create_nested_list_items(doc, items):
            # Different bullet characters for different levels
            bullet_chars = ['•', '◦', '▪', '▫']
            
            # Track item counters for ordered lists at each level
            number_counters = {}
            
            # Track the last level+ordered status to maintain numbering
            last_level_ordered = {}
            
            # Find minimum indentation (base level)
            min_indent = min(item[0] for item in items)
            
            # Process each item
            for indent, content, is_ordered, line_num in items:
                # Calculate actual nesting level (relative to minimum indent)
                level = (indent - min_indent) // 2
                
                # Clean up the content by removing any trailing asterisks or other markdown issues
                if isinstance(content, str):
                    # Fix issues with trailing asterisks in list items
                    content = re.sub(r'\\\*$', '', content.rstrip())  # Remove escaped asterisk at end
                    content = re.sub(r'\*+$', '', content.rstrip())   # Remove trailing asterisks
                    content = content.rstrip()                        # Remove trailing whitespace
                
                # Create paragraph
                p = doc.add_paragraph()
                p.style = 'Normal'
                
                # Visual indentation (just for clarity in the document)
                visual_indent = '  ' * level
                
                # Create a unique key for this level and list type
                level_key = f"{level}_{is_ordered}"
                
                if is_ordered:
                    # For ordered lists, track numbers at each level
                    # Check if we're continuing an existing ordered list at this level
                    if level_key in number_counters:
                        number_counters[level_key] += 1
                    else:
                        # Start a new counter for this level
                        number_counters[level_key] = 1
                    
                    # Different numbering for different levels
                    if level == 0:
                        # First level: 1., 2., 3.
                        prefix = f"{number_counters[level_key]}."
                    elif level == 1:
                        # Second level: a., b., c. 
                        prefix = f"{chr(96 + number_counters[level_key])}."
                    else:
                        # Deeper levels: i., ii., iii.
                        prefix = f"{number_counters[level_key]}."
                    
                    # Add the prefix
                    p.add_run(f"{visual_indent}{prefix} ")
                    
                    # Add the content with formatting
                    add_formatted_text_to_paragraph(p, content)
                    
                    # Update the last level+ordered status
                    last_level_ordered[level] = True
                else:
                    # For unordered lists, use appropriate bullet character
                    bullet = bullet_chars[level % len(bullet_chars)]
                    
                    # Add the bullet
                    p.add_run(f"{visual_indent}{bullet} ")
                    
                    # Add the content with formatting
                    add_formatted_text_to_paragraph(p, content)
                    
                    # Update the last level+ordered status
                    last_level_ordered[level] = False
                    
                    # Reset number counters for this level
                    level_key = f"{level}_True"
                    if level_key in number_counters:
                        del number_counters[level_key]
        
        # Add formatted text to a paragraph
        def add_formatted_text_to_paragraph(paragraph, content):
            if isinstance(content, str):
                # Simple text, just add it
                paragraph.add_run(content)
            else:
                # It's a BeautifulSoup element, process its children
                for child in content.children:
                    if isinstance(child, str):
                        paragraph.add_run(child)
                    elif child.name == 'strong' or child.name == 'b':
                        run = paragraph.add_run(child.get_text())
                        run.bold = True
                    elif child.name == 'em' or child.name == 'i':
                        run = paragraph.add_run(child.get_text())
                        run.italic = True
                    elif child.name == 'code':
                        run = paragraph.add_run(child.get_text())
                        run.font.name = 'Courier New'
                        run.font.size = Pt(10)
                    else:
                        # For other elements, just get their text
                        paragraph.add_run(child.get_text())

        # Helper function to process text with markdown formatting
        def process_text_with_markdown(text):
            # Clean up problematic markdown syntax first
            # Fix issues with inconsistent asterisks and markdown syntax
            text = text.replace('\\*', '§ESCAPED_ASTERISK§')  # Temporarily protect escaped asterisks
            text = re.sub(r'\*{3,}', '**', text)  # Replace any *** or more with **
            text = re.sub(r'(\*{2})(\s*\*{1})|\*\s*\*\*', '**', text)  # Fix ** * or * ** patterns
            text = re.sub(r'(\*{1})(\s*\*{2})|\*\*\s*\*', '**', text)  # Fix * ** or ** * patterns
            text = text.replace('§ESCAPED_ASTERISK§', '*')  # Restore escaped asterisks
            
            # Process the original text to find all markdown formatted sections
            segments = []
            
            # First, handle code backticks to prevent issues with other patterns
            code_pattern = r'`(.*?)`'
            code_segments = []
            last_end = 0
            
            for match in re.finditer(code_pattern, text):
                if match.start() > last_end:
                    code_segments.append(('normal', text[last_end:match.start()]))
                code_segments.append(('code', match.group(1)))
                last_end = match.end()
            
            if last_end < len(text):
                code_segments.append(('normal', text[last_end:]))
            
            # Now process bold and italic in each non-code segment
            for segment_type, segment_text in code_segments:
                if segment_type == 'code':
                    segments.append(('code', segment_text))
                else:
                    # Process bold (must be done before italic to handle nested formatting)
                    bold_pattern = r'\*\*(.*?)\*\*|__(.*?)__'
                    bold_segments = []
                    last_end = 0
                    
                    for match in re.finditer(bold_pattern, segment_text):
                        if match.start() > last_end:
                            bold_segments.append(('normal', segment_text[last_end:match.start()]))
                        
                        # Add the bold text (without ** or __)
                        bold_text = match.group(1) if match.group(1) is not None else match.group(2)
                        bold_segments.append(('bold', bold_text))
                        
                        last_end = match.end()
                    
                    if last_end < len(segment_text):
                        bold_segments.append(('normal', segment_text[last_end:]))
                    
                    # Process italic in each non-bold segment
                    for bold_type, bold_content in bold_segments:
                        if bold_type == 'bold':
                            # For bold text, check for nested italic
                            italic_pattern = r'\*(.*?)\*|_(.*?)_'
                            italic_segments = []
                            last_end = 0
                            
                            for match in re.finditer(italic_pattern, bold_content):
                                if match.start() > last_end:
                                    italic_segments.append(('bold', bold_content[last_end:match.start()]))
                                
                                # Add the italic text (without * or _)
                                italic_text = match.group(1) if match.group(1) is not None else match.group(2)
                                italic_segments.append(('bold+italic', italic_text))
                                
                                last_end = match.end()
                            
                            if last_end < len(bold_content):
                                italic_segments.append(('bold', bold_content[last_end:]))
                            
                            segments.extend(italic_segments)
                        else:
                            # For normal text, check for italic
                            italic_pattern = r'\*(.*?)\*|_(.*?)_'
                            italic_segments = []
                            last_end = 0
                            
                            for match in re.finditer(italic_pattern, bold_content):
                                if match.start() > last_end:
                                    italic_segments.append(('normal', bold_content[last_end:match.start()]))
                                
                                # Add the italic text (without * or _)
                                italic_text = match.group(1) if match.group(1) is not None else match.group(2)
                                italic_segments.append(('italic', italic_text))
                                
                                last_end = match.end()
                            
                            if last_end < len(bold_content):
                                italic_segments.append(('normal', bold_content[last_end:]))
                            
                            segments.extend(italic_segments)
            
            return segments

        # Find all list blocks in the markdown
        list_blocks = identify_list_blocks(markdown_lines)
        
        # Create a map to find list blocks by line
        list_block_map = {}
        for start, end in list_blocks:
            for i in range(start, end + 1):
                list_block_map[i] = (start, end)
        
        # Keep track of whether we've seen the H1 heading
        h1_heading_processed = False
        
        # Track lines that have been processed
        processed_lines = set()
        
        # Track processed list blocks to avoid duplicates
        processed_list_blocks = set()
        
        # Create an additional set to track which paragraphs have inline lists processed
        inline_list_paragraphs = set()
        
        # ===== PROCESS THE DOCUMENT LINE BY LINE TO MAINTAIN ORDER =====
        i = 0
        while i < len(markdown_lines):
            line = markdown_lines[i].strip()
            
            # Skip empty lines and horizontal rules (---, ___, ***)
            if not line or re.match(r'^[-_*]{3,}$', line):
                processed_lines.add(i)
                i += 1
                continue
            
            # Check if it's a list block
            if i in list_block_map:
                start, end = list_block_map[i]
                
                # Skip if we already processed this list block
                if start not in processed_list_blocks:
                    process_list_block(doc, markdown_lines, start, end)
                    processed_list_blocks.add((start, end))
                    
                    # Mark all lines in this block as processed
                    for j in range(start, end + 1):
                        processed_lines.add(j)
                
                # Skip to the end of the block
                i = end + 1
                continue
            
            # Check if it's a heading
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if heading_match:
                level = len(heading_match.group(1))
                text = heading_match.group(2).strip()
                
                # Adjust heading levels
                if level == 1:
                    # First # becomes Heading 0
                    style_name = 'Heading 0'
                else:
                    # Subsequent levels are shifted down by 1
                    style_name = f'Heading {level - 1}'
                
                # Create heading paragraph
                p = doc.add_paragraph()
                try:
                    p.style = style_name
                except KeyError:
                    print(f"Warning: Style '{style_name}' not found. Using 'Normal'.")
                    p.style = 'Normal'
                
                # Clean any manual numbering
                text = clean_heading_text(text)
                
                # Process any markdown formatting in headings
                segments = process_text_with_markdown(text)
                for format_type, segment_text in segments:
                    run = p.add_run(segment_text)
                    if format_type == 'bold' or format_type == 'bold+italic':
                        run.bold = True
                    if format_type == 'italic' or format_type == 'bold+italic':
                        run.italic = True
                    if format_type == 'code':
                        run.font.name = 'Courier New'
                        run.font.size = Pt(10)
                
                processed_lines.add(i)
                i += 1
                continue
                
                # Create heading paragraph
                p = doc.add_paragraph()
                try:
                    p.style = f'Heading {level}'
                except KeyError:
                    print(f"Warning: Style 'Heading {level}' not found. Using 'Normal'.")
                    p.style = 'Normal'
                
                # Clean any manual numbering
                text = clean_heading_text(text)
                
                # Process any markdown formatting in headings
                segments = process_text_with_markdown(text)
                for format_type, segment_text in segments:
                    run = p.add_run(segment_text)
                    if format_type == 'bold' or format_type == 'bold+italic':
                        run.bold = True
                    if format_type == 'italic' or format_type == 'bold+italic':
                        run.italic = True
                    if format_type == 'code':
                        run.font.name = 'Courier New'
                        run.font.size = Pt(10)
                
                processed_lines.add(i)
                i += 1
                continue
            
            # Check if it's a code block
            if line.startswith('```'):
                # Extract language if specified
                language = line[3:].strip()
                
                # Collect code lines until closing ```
                code_lines = []
                i += 1
                start_code_block = i - 1
                
                while i < len(markdown_lines) and not markdown_lines[i].strip().startswith('```'):
                    code_lines.append(markdown_lines[i])
                    processed_lines.add(i)
                    i += 1
                
                # Skip the closing ```
                end_code_block = i
                processed_lines.add(i)
                i += 1
                
                # Add code block
                if code_lines:
                    code_text = '\n'.join(code_lines)
                    add_code_block(doc, code_text, language)
                
                continue
            
            # If it's a table, let the HTML processor handle it
            if line.startswith('|') and i + 1 < len(markdown_lines) and markdown_lines[i + 1].strip().startswith('|'):
                # Find table in HTML
                tables = soup.find_all('table')
                table_start = i
                
                # Skip ahead to find the end of the table
                while i < len(markdown_lines) and (markdown_lines[i].strip().startswith('|') or markdown_lines[i].strip() == ''):
                    processed_lines.add(i)
                    i += 1
                table_end = i - 1
                
                # Process the table only once
                for table in tables:
                    # Add the table if it's not been processed
                    table_text = ''.join(markdown_lines[table_start:table_end+1])
                    add_table(doc, table)
                    break
                    
                continue
            
            # Only process each line once to avoid duplication
            if i in processed_lines:
                i += 1
                continue
                
            # It's a regular paragraph, find in HTML to preserve formatting
            paragraphs = soup.find_all('p')
            plain_text_map = {}
            for p in paragraphs:
                plain_text = p.get_text().strip()
                plain_text_map[plain_text] = p
            
            # Check if the line looks like an inline list (contains backslash followed by dash/asterisk or ends with backslash)
            is_inline_list = re.search(r'\\(\s*[-*+])', line) is not None or line.endswith('\\')
            
            # Check if this is part of the "Text Formatting" section
            in_text_formatting_section = False
            for j in range(max(0, i-20), i):
                if 'Text Formatting' in markdown_lines[j]:
                    in_text_formatting_section = True
                    break
            
            # Try to find closest match
            matched = False
            for plain_text, p_element in plain_text_map.items():
                if line in plain_text or plain_text in line:
                    # Check if this paragraph contains a list item that will be processed separately
                    is_list_item = False
                    for list_start, list_end in list_blocks:
                        for list_line in range(list_start, list_end + 1):
                            list_text = markdown_lines[list_line].strip()
                            if list_text and (list_text in plain_text or plain_text in list_text):
                                is_list_item = True
                                break
                        if is_list_item:
                            break
                
                    # Special handling for inline lists with backslashes or text formatting section
                    if is_inline_list or in_text_formatting_section:
                        # Skip paragraph if we're in the text formatting section and this might be a duplicate
                        if in_text_formatting_section:
                            text_content = p_element.get_text().strip()
                            for future_line in range(i+1, min(i+15, len(markdown_lines))):
                                future_text = markdown_lines[future_line].strip()
                                # If there's a bullet point version of this content coming up soon, skip this paragraph
                                if future_text.startswith('•') and text_content.endswith(future_text[1:].strip()):
                                    matched = True
                                    break
                                # Also detect if content has "Bold text is supported" or similar patterns
                                if "Bold text is supported" in future_text and "Bold text is supported" in text_content:
                                    matched = True
                                    break
                                if "Italic text is supported" in future_text and "Italic text is supported" in text_content:
                                    matched = True
                                    break
                        
                        # For normal inline lists, check similarity with existing list blocks
                        else:
                            has_matching_list = False
                            for start_end in processed_list_blocks:
                                list_start, list_end = start_end
                                list_content = ' '.join([markdown_lines[j].strip() for j in range(list_start, list_end+1)])
                                p_content = p_element.get_text().strip()
                                
                                # If the list content is similar to paragraph content, skip paragraph
                                if len(list_content) > 0 and len(p_content) > 0:
                                    similarity = len(set(list_content.split()) & set(p_content.split())) / len(set(p_content.split()))
                                    if similarity > 0.7:  # If 70% of words match
                                        has_matching_list = True
                                        break
                            
                            if not has_matching_list and not in_text_formatting_section:
                                add_paragraph(doc, p_element, 'Normal')
                                matched = True
                    elif not is_list_item:
                        add_paragraph(doc, p_element, 'Normal')
                        matched = True
                    break
            
            if not matched:
                # Skip the line if it appears to be a formatting artifact 
                # like a single dash or lone asterisk, etc.
                if line == '--' or line == '*' or line == '---':
                    processed_lines.add(i)
                    
                    # For horizontal rules, add a proper horizontal line
                    if line == '---':
                        hr_paragraph = doc.add_paragraph()
                        hr_paragraph.style = 'Normal'
                        # Add a thin horizontal line by adding a bottom border to the paragraph
                        hr_paragraph.paragraph_format.bottom_border.width = Pt(1)
                        
                    i += 1
                    continue
                    
                # Special handling for contact information at the end of the document
                if line.startswith('*For') and line.endswith('*'):
                    p = doc.add_paragraph()
                    p.style = 'Normal'
                    run = p.add_run(line.strip('*'))
                    run.italic = True
                    processed_lines.add(i)
                    i += 1
                    continue
                    
                # Skip line if it appears to be part of a list that will be processed separately
                is_list_line = False
                for list_start, list_end in list_blocks:
                    if list_start <= i <= list_end:
                        is_list_line = True
                        break
                
                if not is_list_line:
                    # If no match found, try to process markdown formatting directly
                    p = doc.add_paragraph()
                    p.style = 'Normal'
                    
                    # Fix special cases for contact info/footnotes
                    if line.startswith('*For') and line.endswith('*'):
                        # Handle the contact info line as italic
                        run = p.add_run(line.strip('*'))
                        run.italic = True
                    else:
                        # Process markdown formatting
                        segments = process_text_with_markdown(line)
                        for format_type, segment_text in segments:
                            run = p.add_run(segment_text)
                            if format_type == 'bold' or format_type == 'bold+italic':
                                run.bold = True
                            if format_type == 'italic' or format_type == 'bold+italic':
                                run.italic = True
                            if format_type == 'code':
                                run.font.name = 'Courier New'
                                run.font.size = Pt(10)
            
            processed_lines.add(i)
            i += 1
        
        # Save the document
        doc.save(output_path)
        print(f"Document saved to {output_path}")
        return True
        
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        return False

def convert_file(md_file_path, template_path, output_path=None):
    """
    Convert a single Markdown file to DOCX
    
    Args:
        md_file_path (str): Path to the markdown file
        template_path (str): Path to the Word template
        output_path (str, optional): Path to save the output file. If None, will use the same name as the input file.
    
    Returns:
        tuple: (success, output_path)
    """
    try:
        # If output_path is not specified, create one based on input file
        if output_path is None:
            output_dir = os.path.dirname(md_file_path)
            file_name = os.path.splitext(os.path.basename(md_file_path))[0]
            output_path = os.path.join(output_dir, f"{file_name}.docx")
        
        # If output_path is a directory, append file name
        elif os.path.isdir(output_path):
            file_name = os.path.splitext(os.path.basename(md_file_path))[0]
            output_path = os.path.join(output_path, f"{file_name}.docx")
            
        # Create parent directory if it doesn't exist
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        
        # Call conversion function
        success = create_docx_from_readme(md_file_path, template_path, output_path)
        
        return success, output_path
    except Exception as e:
        import traceback
        print(f"Error converting file: {str(e)}")
        print(traceback.format_exc())
        return False, None

def convert_folder(folder_path, template_path, output_folder=None):
    """
    Convert all Markdown files in a folder to DOCX
    
    Args:
        folder_path (str): Path to the folder containing markdown files
        template_path (str): Path to the Word template
        output_folder (str, optional): Path to save the output files. If None, will use the same folder as the input.
    
    Returns:
        list: List of (success, output_path) tuples for each converted file
    """
    try:
        # If output_folder is not specified, use the input folder
        if output_folder is None:
            output_folder = folder_path
        
        # Make sure the output folder exists
        os.makedirs(output_folder, exist_ok=True)
        
        results = []
        
        # Find all markdown files in the folder
        md_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.md')]
        
        # Convert each file
        for md_file in md_files:
            md_file_path = os.path.join(folder_path, md_file)
            file_name = os.path.splitext(md_file)[0]
            output_path = os.path.join(output_folder, f"{file_name}.docx")
            
            # Call the convert_file function to handle the conversion
            success, actual_output_path = convert_file(md_file_path, template_path, output_path)
            results.append((success, actual_output_path if success else output_path))
        
        return results
    except Exception as e:
        import traceback
        print(f"Error converting folder: {str(e)}")
        print(traceback.format_exc())
        return [(False, None)]
