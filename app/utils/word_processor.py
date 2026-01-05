from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX
import os

class WordProcessor:
    def __init__(self):
        pass
    
    def extract_content(self, file_path):
        """
        Extract content from a Word document (.docx or .doc)
        Returns a dictionary with text, images, and formatting information
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        try:
            doc = Document(file_path)
            
            content = {
                'paragraphs': [],
                'tables': [],
                'images': [],
                'metadata': {
                    'title': '',
                    'author': '',
                    'subject': '',
                    'keywords': ''
                }
            }
            
            # Extract paragraphs
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    para_data = {
                        'text': paragraph.text,
                        'style': paragraph.style.name if paragraph.style else 'Normal',
                        'alignment': str(paragraph.alignment),
                        'runs': []
                    }
                    
                    # Extract run formatting
                    for run in paragraph.runs:
                        run_data = {
                            'text': run.text,
                            'bold': run.bold,
                            'italic': run.italic,
                            'underline': run.underline,
                            'font_size': run.font.size.pt if run.font.size else None,
                            'font_name': run.font.name if run.font.name else None
                        }
                        para_data['runs'].append(run_data)
                    
                    content['paragraphs'].append(para_data)
            
            # Extract tables
            for table in doc.tables:
                table_data = {
                    'rows': []
                }
                
                for row in table.rows:
                    row_data = {
                        'cells': []
                    }
                    
                    for cell in row.cells:
                        cell_data = {
                            'text': cell.text,
                            'paragraphs': []
                        }
                        
                        for paragraph in cell.paragraphs:
                            cell_data['paragraphs'].append({
                                'text': paragraph.text,
                                'style': paragraph.style.name if paragraph.style else 'Normal'
                            })
                        
                        row_data['cells'].append(cell_data)
                    
                    table_data['rows'].append(row_data)
                
                content['tables'].append(table_data)
            
            # Extract metadata if available
            if hasattr(doc.core_properties, 'title') and doc.core_properties.title:
                content['metadata']['title'] = doc.core_properties.title
            if hasattr(doc.core_properties, 'author') and doc.core_properties.author:
                content['metadata']['author'] = doc.core_properties.author
            if hasattr(doc.core_properties, 'subject') and doc.core_properties.subject:
                content['metadata']['subject'] = doc.core_properties.subject
            if hasattr(doc.core_properties, 'keywords') and doc.core_properties.keywords:
                content['metadata']['keywords'] = doc.core_properties.keywords
            
            return content
            
        except Exception as e:
            raise Exception(f"Error processing Word document: {str(e)}")
    
    def get_document_info(self, file_path):
        """
        Get basic information about the Word document
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        try:
            doc = Document(file_path)
            
            info = {
                'paragraph_count': len(doc.paragraphs),
                'table_count': len(doc.tables),
                'section_count': len(doc.sections),
                'file_size': os.path.getsize(file_path)
            }
            
            return info
            
        except Exception as e:
            raise Exception(f"Error getting document info: {str(e)}")
    
    def _remove_highlighting(self, run):
        """
        Remove highlighting from a run by both API and XML methods.
        This ensures highlighting is completely removed.
        """
        try:
            # Method 1: Use API
            if hasattr(run.font, 'highlight_color'):
                run.font.highlight_color = None
            # Method 2: Remove from XML directly - more thorough approach
            if hasattr(run, '_element'):
                rPr = run._element.get_or_add_rPr()
                if rPr is not None:
                    # Find all highlight elements (there might be multiple or nested)
                    namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                    # Method 1: Find by exact namespace
                    highlights = rPr.findall(f'{namespace}highlight')
                    for highlight in highlights:
                        rPr.remove(highlight)
                    # Method 2: Find by tag name (in case namespace is different)
                    for elem in list(rPr):
                        if elem.tag.endswith('}highlight') or 'highlight' in elem.tag.lower():
                            rPr.remove(elem)
                    # Method 3: Try to set the attribute directly if it exists
                    if hasattr(rPr, 'highlight'):
                        rPr.highlight = None
                    # Method 4: Also check for highlight attribute on the element itself
                    if hasattr(rPr, 'get'):
                        highlight_attr = rPr.get(f'{namespace}highlight')
                        if highlight_attr is not None:
                            rPr.remove(highlight_attr)
        except Exception:
            pass  # If highlighting can't be removed, continue
    
    def _remove_highlighting_from_all_runs(self, paragraph, start_idx, end_idx):
        """
        Remove highlighting from all runs between start_idx and end_idx (inclusive).
        This ensures we remove highlighting from all runs that were part of the placeholder.
        """
        for i in range(start_idx, min(end_idx + 1, len(paragraph.runs))):
            try:
                self._remove_highlighting(paragraph.runs[i])
            except Exception:
                pass
    
    def _remove_all_highlighting_from_paragraph(self, paragraph):
        """
        Remove highlighting from all runs in a paragraph.
        This is a comprehensive cleanup method to ensure no highlighting remains.
        """
        for run in paragraph.runs:
            try:
                self._remove_highlighting(run)
            except Exception:
                pass
    
    def _replace_placeholder_in_paragraph(self, paragraph, placeholder, value):
        """
        Replace a placeholder in a paragraph, handling cases where the placeholder
        may be split across multiple runs (e.g., due to formatting like bold).
        """
        placeholder_text = f'{{{placeholder}}}'
        
        # Check if placeholder exists in paragraph text
        if placeholder_text not in paragraph.text:
            return False
        
        # Convert value to string, handling None and empty values
        replacement_value = str(value) if value is not None else ''
        
        # First, try simple replacement if placeholder is in a single run
        for run in paragraph.runs:
            if placeholder_text in run.text:
                # Remove highlighting BEFORE replacement to ensure it's gone
                self._remove_highlighting(run)
                # Explicitly set highlight_color to None before text replacement
                if hasattr(run.font, 'highlight_color'):
                    run.font.highlight_color = None
                # Replace the placeholder with the value
                run.text = run.text.replace(placeholder_text, replacement_value)
                # Remove highlighting again AFTER replacement to be thorough
                self._remove_highlighting(run)
                # One more explicit check after text change
                if hasattr(run.font, 'highlight_color'):
                    run.font.highlight_color = None
                return True
        
        # Placeholder is split across multiple runs - need to handle this
        # Strategy: Replace text in the paragraph by working with runs
        full_text = paragraph.text
        placeholder_start = full_text.find(placeholder_text)
        
        if placeholder_start == -1:
            return False
        
        placeholder_end = placeholder_start + len(placeholder_text)
        
        # Find which runs contain parts of the placeholder
        current_pos = 0
        start_run_idx = None
        end_run_idx = None
        
        for i, run in enumerate(paragraph.runs):
            run_length = len(run.text)
            run_start = current_pos
            run_end = current_pos + run_length
            
            if start_run_idx is None and run_start <= placeholder_start < run_end:
                start_run_idx = i
            if run_start < placeholder_end <= run_end:
                end_run_idx = i
                break
            
            current_pos = run_end
        
        if start_run_idx is None or end_run_idx is None:
            return False
        
        # Convert value to string, handling None and empty values
        replacement_value = str(value) if value is not None else ''
        
        # Safety check: if start and end are the same, it should have been caught above
        # But handle it just in case
        if start_run_idx == end_run_idx:
            # Placeholder should be in a single run - try simple replacement
            if start_run_idx < len(paragraph.runs):
                # Remove highlighting before replacement
                self._remove_highlighting(paragraph.runs[start_run_idx])
                if hasattr(paragraph.runs[start_run_idx].font, 'highlight_color'):
                    paragraph.runs[start_run_idx].font.highlight_color = None
                # Replace the placeholder
                paragraph.runs[start_run_idx].text = paragraph.runs[start_run_idx].text.replace(placeholder_text, replacement_value)
                # Remove highlighting/background color after replacement
                self._remove_highlighting(paragraph.runs[start_run_idx])
                if hasattr(paragraph.runs[start_run_idx].font, 'highlight_color'):
                    paragraph.runs[start_run_idx].font.highlight_color = None
                return True
            return False
        
        # Calculate text before and after placeholder in the full text
        text_before = full_text[:placeholder_start]
        text_after = full_text[placeholder_end:]
        
        # Now we need to reconstruct the paragraph
        # Save the original paragraph text with replacement
        new_paragraph_text = text_before + replacement_value + text_after
        
        # Calculate positions within individual runs
        pos = 0
        start_run_text_before = ''
        end_run_text_after = ''
        
        for i, run in enumerate(paragraph.runs):
            run_text = run.text
            run_len = len(run_text)
            
            if i == start_run_idx:
                # Calculate how much of this run is before the placeholder
                offset_in_run = placeholder_start - pos
                start_run_text_before = run_text[:offset_in_run]
            elif i == end_run_idx:
                # Calculate how much of this run is after the placeholder
                offset_in_run = placeholder_end - pos
                end_run_text_after = run_text[offset_in_run:]
                break
            
            pos += run_len
        
        # Remove highlighting from all runs that contained the placeholder BEFORE replacement
        # This ensures we catch highlighting that might be on any part of the placeholder
        self._remove_highlighting_from_all_runs(paragraph, start_run_idx, end_run_idx)
        
        # Replace the placeholder: update start run, remove middle runs, update/remove end run
        if start_run_idx < len(paragraph.runs):
            # Remove highlighting BEFORE updating text
            self._remove_highlighting(paragraph.runs[start_run_idx])
            # Explicitly set highlight_color to None before text replacement
            if hasattr(paragraph.runs[start_run_idx].font, 'highlight_color'):
                paragraph.runs[start_run_idx].font.highlight_color = None
            # Update start run with text before + replacement value
            # Ensure we're using replacement_value (handles None/empty properly)
            new_text = start_run_text_before + replacement_value
            paragraph.runs[start_run_idx].text = new_text
            # Verify the text was set correctly (safeguard)
            if paragraph.runs[start_run_idx].text != new_text:
                # If text wasn't set correctly, try setting it again
                paragraph.runs[start_run_idx].text = new_text
            # Remove highlighting/background color from the run (again, to be sure)
            self._remove_highlighting(paragraph.runs[start_run_idx])
            # One more explicit check after text change
            if hasattr(paragraph.runs[start_run_idx].font, 'highlight_color'):
                paragraph.runs[start_run_idx].font.highlight_color = None
        
        # Handle end run - save formatting BEFORE removing runs (important!)
        orig_end_run_formatting = None
        orig_end_run_highlight = None
        if end_run_idx < len(paragraph.runs):
            orig_end_run = paragraph.runs[end_run_idx]
            orig_end_run_formatting = {
                'bold': orig_end_run.bold,
                'italic': orig_end_run.italic,
                'font_size': orig_end_run.font.size
            }
            orig_end_run_highlight = orig_end_run.font.highlight_color
        
        # Remove middle runs (between start and end, exclusive)
        runs_to_remove = list(range(start_run_idx + 1, end_run_idx))
        for i in reversed(runs_to_remove):
            if i < len(paragraph.runs):
                paragraph._element.remove(paragraph.runs[i]._element)
        
        # Handle end run after removals
        if end_run_idx > start_run_idx:
            # After removing middle runs, the end_run_idx has shifted
            remaining_runs_count = len(paragraph.runs)
            expected_end_idx = end_run_idx - len(runs_to_remove)
            
            # Ensure we have valid indices
            if expected_end_idx < remaining_runs_count and expected_end_idx >= start_run_idx:
                # Check if the end run still exists and is different from start run
                if expected_end_idx != start_run_idx and expected_end_idx < len(paragraph.runs):
                    # Remove highlighting BEFORE updating text
                    self._remove_highlighting(paragraph.runs[expected_end_idx])
                    # Explicitly set highlight_color to None before text replacement
                    if hasattr(paragraph.runs[expected_end_idx].font, 'highlight_color'):
                        paragraph.runs[expected_end_idx].font.highlight_color = None
                    # Update the end run with remaining text
                    paragraph.runs[expected_end_idx].text = end_run_text_after
                    # Remove highlighting if this run had part of the placeholder
                    self._remove_highlighting(paragraph.runs[expected_end_idx])
                    # One more explicit check after text change
                    if hasattr(paragraph.runs[expected_end_idx].font, 'highlight_color'):
                        paragraph.runs[expected_end_idx].font.highlight_color = None
                elif end_run_text_after:
                    # End run was merged or removed, need to add a new run for remaining text
                    # Use formatting from the original end run if we saved it
                    new_run = paragraph.add_run(end_run_text_after)
                    if orig_end_run_formatting:
                        if orig_end_run_formatting['bold'] is not None:
                            new_run.bold = orig_end_run_formatting['bold']
                        if orig_end_run_formatting['italic'] is not None:
                            new_run.italic = orig_end_run_formatting['italic']
                        if orig_end_run_formatting['font_size'] is not None:
                            new_run.font.size = orig_end_run_formatting['font_size']
                    # Ensure no highlighting on new run - explicitly set to None
                    if hasattr(new_run.font, 'highlight_color'):
                        new_run.font.highlight_color = None
                    self._remove_highlighting(new_run)
            elif end_run_text_after:
                # End run index is invalid, add a new run for remaining text
                new_run = paragraph.add_run(end_run_text_after)
                if orig_end_run_formatting:
                    if orig_end_run_formatting['bold'] is not None:
                        new_run.bold = orig_end_run_formatting['bold']
                    if orig_end_run_formatting['italic'] is not None:
                        new_run.italic = orig_end_run_formatting['italic']
                    if orig_end_run_formatting['font_size'] is not None:
                        new_run.font.size = orig_end_run_formatting['font_size']
                # Ensure no highlighting on new run - explicitly set to None
                if hasattr(new_run.font, 'highlight_color'):
                    new_run.font.highlight_color = None
                self._remove_highlighting(new_run)
        
        return True
    
    def fill_placeholders(self, template_path, output_path, data):
        """
        Fill placeholders in the format {FieldName} in the Word template with values from data dict.
        Save the filled document to output_path.
        Handles placeholders that may be split across multiple runs due to formatting.
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")
        doc = Document(template_path)
        
        # Replace in paragraphs
        for paragraph in doc.paragraphs:
            for key, value in data.items():
                self._replace_placeholder_in_paragraph(paragraph, key, value)
            # Final cleanup: Remove any remaining highlighting from all runs in this paragraph
            self._remove_all_highlighting_from_paragraph(paragraph)
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for key, value in data.items():
                            self._replace_placeholder_in_paragraph(paragraph, key, value)
                        # Final cleanup: Remove any remaining highlighting from all runs in this paragraph
                        self._remove_all_highlighting_from_paragraph(paragraph)
        
        # Final comprehensive pass: Remove highlighting from ALL paragraphs and runs
        # This ensures no highlighting remains anywhere in the document, even if it wasn't part of a placeholder
        for paragraph in doc.paragraphs:
            self._remove_all_highlighting_from_paragraph(paragraph)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._remove_all_highlighting_from_paragraph(paragraph)
        
        doc.save(output_path) 