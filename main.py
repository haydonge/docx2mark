import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from markitdown import MarkItDown
import os
import subprocess
from datetime import datetime
from PIL import Image
import hashlib

class MarkItDownGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MarkItDown Converter")
        self.root.geometry("600x400")
        
        # Initialize converter
        self.converter = MarkItDown()
        
        style = ttk.Style()
        style.configure('TButton', padding=6)
        style.configure('TLabel', padding=6)
        
        main_frame = ttk.Frame(root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.file_path = tk.StringVar()
        ttk.Label(main_frame, text="Select a file to convert:").grid(row=0, column=0, sticky=tk.W)
        
        file_entry = ttk.Entry(main_frame, textvariable=self.file_path, width=50)
        file_entry.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W+tk.E)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=1, padx=5, pady=5)
        
        browse_btn = ttk.Button(button_frame, text="Browse File", command=self.browse_file)
        browse_btn.pack(side=tk.LEFT, padx=2)
        
        browse_dir_btn = ttk.Button(button_frame, text="Browse Directory", command=self.browse_directory)
        browse_dir_btn.pack(side=tk.LEFT, padx=2)
        
        convert_btn = ttk.Button(main_frame, text="Convert to Markdown", command=self.convert_file)
        convert_btn.grid(row=2, column=0, columnspan=2, pady=20)
        
        ttk.Label(main_frame, text="Conversion Output:").grid(row=3, column=0, sticky=tk.W)
        
        self.output_text = tk.Text(main_frame, height=12, wrap=tk.WORD)
        self.output_text.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.output_text.yview)
        scrollbar.grid(row=4, column=2, sticky=(tk.N, tk.S))
        self.output_text['yscrollcommand'] = scrollbar.set
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path.set(file_path)
            
    def browse_directory(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.file_path.set(dir_path)
            
    def convert_single_file(self, input_file):
        try:
            result = self.converter.convert(input_file)
            output_path = os.path.splitext(input_file)[0] + '.md'
            
            with open(output_path, 'w', encoding='utf-8') as md_file:
                md_file.write(result.text_content)
                
            return output_path
        except Exception as e:
            raise Exception(f"Error converting {os.path.basename(input_file)}: {str(e)}")
            
    def convert_file(self):
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showerror("Error", "Please select a file or directory first.")
            return

        try:
            self.output_text.delete(1.0, tk.END)  # Clear previous output
            
            if os.path.isfile(file_path):
                # Single file conversion
                if file_path.lower().endswith('.docx'):
                    output_path = self.convert_single_file(file_path)
                    self.output_text.insert(tk.END, f"Successfully converted {file_path} to {output_path}\n")
                else:
                    self.output_text.insert(tk.END, f"Skipped {file_path} - not a DOCX file\n")
            else:
                # Directory conversion
                found_files = False
                for root, _, files in os.walk(file_path):
                    for file in files:
                        if file.lower().endswith('.docx'):
                            found_files = True
                            file_path = os.path.join(root, file)
                            try:
                                output_path = self.convert_single_file(file_path)
                                self.output_text.insert(tk.END, f"Successfully converted {file_path} to {output_path}\n")
                            except Exception as e:
                                self.output_text.insert(tk.END, f"Error converting {file_path}: {str(e)}\n")
                
                if not found_files:
                    self.output_text.insert(tk.END, "No DOCX files found in the selected directory.\n")
            
            self.output_text.see(tk.END)
            messagebox.showinfo("Success", "Conversion process completed!")
        except Exception as e:
            self.output_text.insert(tk.END, f"Error: {str(e)}\n")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

def main():
    root = tk.Tk()
    app = MarkItDownGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()