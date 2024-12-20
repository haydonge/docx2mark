import os
import glob
from datetime import datetime

def add_frontmatter(file_path):
    # Get the filename without extension
    filename = os.path.basename(file_path)
    base_name = os.path.splitext(filename)[0]
    
    # Extract document number for slug
    doc_number = ''.join(filter(str.isdigit, base_name))
    
    # Determine category
    category = "Sperm" if "male" in base_name.lower() else "eggdonor"
    
    # Create frontmatter
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
    
    # Read the existing content
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
    
    # Write the frontmatter and content back
    with open(file_path, 'w', encoding='utf-8') as file:
        file.write(frontmatter + content)

def main():
    # Process all markdown files in the current directory
    md_files = glob.glob('*.md')
    for md_file in md_files:
        if md_file != 'README.md':  # Skip README if exists
            add_frontmatter(md_file)

if __name__ == "__main__":
    main()
