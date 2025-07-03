from docx import Document
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
    
    def fill_placeholders(self, template_path, output_path, data):
        """
        Fill placeholders in the format {FieldName} in the Word template with values from data dict.
        Save the filled document to output_path.
        """
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"Template not found: {template_path}")
        doc = Document(template_path)
        # Replace in paragraphs
        for paragraph in doc.paragraphs:
            for key, value in data.items():
                if f'{{{key}}}' in paragraph.text:
                    inline = paragraph.runs
                    for i in range(len(inline)):
                        if f'{{{key}}}' in inline[i].text:
                            inline[i].text = inline[i].text.replace(f'{{{key}}}', str(value))
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for key, value in data.items():
                            if f'{{{key}}}' in paragraph.text:
                                inline = paragraph.runs
                                for i in range(len(inline)):
                                    if f'{{{key}}}' in inline[i].text:
                                        inline[i].text = inline[i].text.replace(f'{{{key}}}', str(value))
        doc.save(output_path) 