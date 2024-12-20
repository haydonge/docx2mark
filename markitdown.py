import os
import mammoth
import re
from datetime import datetime
import hashlib

class MarkItDown:
    def __init__(self):
        self.image_dir = "media"
        self.image_hashes = {}

    def convert(self, input_file):
        """Convert the input file to markdown format"""
        # Create media directory if it doesn't exist
        if not os.path.exists(self.image_dir):
            os.makedirs(self.image_dir)
            
        # Store current docx path for image handling
        self.current_docx = input_file
        self.image_hashes.clear()

        # Convert DOCX to markdown
        with open(input_file, 'rb') as docx_file:
            result = mammoth.convert_to_markdown(
                docx_file,
                convert_image=mammoth.images.img_element(self.handle_image)
            )

        # Process the markdown content
        content = self._process_markdown(result.value)
        
        # Add frontmatter
        content = self._add_frontmatter(content, input_file)
        
        return type('ConversionResult', (), {
            'text_content': content,
            'messages': result.messages
        })

    def _process_markdown(self, content):
        """Process and clean the markdown content"""
        # Remove base64 images
        content = re.sub(r'!\[\]\(data:image/png;base64,[^\)]+\)', '', content)
        
        # Convert text to table format
        lines = content.split('\n')
        fixed_lines = []
        current_row = []
        
        for line in lines:
            line = line.strip()
            if not line:
                if current_row:
                    # Add the current row as a table
                    fixed_lines.extend(self._format_table_row(current_row))
                    current_row = []
                fixed_lines.append('')
                continue

            # Check if this is a header line
            if line.startswith('__') and line.endswith('__'):
                if current_row:
                    fixed_lines.extend(self._format_table_row(current_row))
                    current_row = []
                fixed_lines.append(line)
                continue

            # Handle label-value pairs
            if ': ' in line:
                if current_row:
                    fixed_lines.extend(self._format_table_row(current_row))
                current_row = [line]
            elif current_row:
                current_row.append(line)
            else:
                fixed_lines.append(line)

        # Add any remaining row
        if current_row:
            fixed_lines.extend(self._format_table_row(current_row))

        return '\n\n'.join(line for line in fixed_lines if line.strip())

    def _format_table_row(self, row):
        """Format a row as a markdown table"""
        if len(row) != 2:
            return row
        return [
            f'| {row[0]} |',
            '| --- |',
            f'| {row[1]} |',
            ''
        ]

    def _add_frontmatter(self, content, input_file):
        """Add frontmatter to the markdown content"""
        base_name = os.path.splitext(os.path.basename(input_file))[0]
        doc_number = ''.join(filter(str.isdigit, base_name))
        category = "Sperm" if "male" in base_name.lower() else "eggdonor"
        
        frontmatter = f"""---
title: {base_name}
date: 2024/12/20
type: Post
status: Published
slug: donor{doc_number}
tags: donor
summary: 
category: {category}
---

"""
        return frontmatter + content

    def handle_image(self, image):
        """Handle image conversion and storage"""
        with image.open() as image_bytes:
            image_data = image_bytes.read()
            image_hash = hashlib.md5(image_data).hexdigest()[:8]
            
            # Generate filename using document hash and index
            doc_id = hashlib.md5(os.path.basename(self.current_docx).encode()).hexdigest()[:8]
            idx = len(self.image_hashes) + 1
            timestamp = datetime.now().strftime('%H%M%S')
            
            # Get original extension from content type
            extension = image.content_type.split('/')[-1]
            if extension == 'jpeg':
                extension = 'png'
                
            filename = f"{doc_id}_{idx:02d}_{timestamp}.{extension}"
            path = os.path.join(self.image_dir, filename)
            
            # Save image only if it's not a duplicate
            if image_hash not in self.image_hashes:
                with open(path, 'wb') as f:
                    f.write(image_data)
                self.image_hashes[image_hash] = filename
            
            # Return markdown image reference
            return {"src": f"media/{self.image_hashes[image_hash]}"}
