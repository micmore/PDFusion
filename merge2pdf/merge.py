import tkinter as tk
from tkinter import filedialog, messagebox, Listbox
from PIL import Image
from fpdf import FPDF
import os
from docx import Document

class FileToPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File to PDF Converter")
        
        self.upload_button = tk.Button(root, text="Add File", command=self.add_file)
        self.upload_button.pack(pady=10)
        
        self.file_listbox = Listbox(root, selectmode=tk.BROWSE, width=50)
        self.file_listbox.pack(pady=10)
        
        self.remove_button = tk.Button(root, text="Remove Selected File", command=self.remove_file, state=tk.DISABLED)
        self.remove_button.pack(pady=5)
        
        self.merge_button = tk.Button(root, text="Merge to PDF", command=self.merge_to_pdf, state=tk.DISABLED)
        self.merge_button.pack(pady=10)
        
        self.files = []

    def add_file(self):
        file = filedialog.askopenfilename(
            title="Select a File", 
            filetypes=[("Supported Files", "*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.txt;*.docx")]
        )
        
        if file:
            if len(self.files) >= 30:
                messagebox.showwarning("Too many files", "You can only upload up to 30 files.")
            else:
                self.files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
                self.merge_button.config(state=tk.NORMAL)
                self.remove_button.config(state=tk.NORMAL)

    def remove_file(self):
        selected_index = self.file_listbox.curselection()
        if selected_index:
            index = selected_index[0]
            self.file_listbox.delete(index)
            del self.files[index]
            
            # Disable buttons if no files left
            if not self.files:
                self.merge_button.config(state=tk.DISABLED)
                self.remove_button.config(state=tk.DISABLED)

    def convert_text_to_pdf(self, text_file, pdf):
        with open(text_file, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            for line in lines:
                pdf.add_page()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.set_font("Arial", size=12)
                pdf.multi_cell(0, 10, line)

    def convert_docx_to_pdf(self, docx_file, pdf):
        doc = Document(docx_file)
        for paragraph in doc.paragraphs:
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, paragraph.text)

    def merge_to_pdf(self):
        if not self.files:
            messagebox.showerror("No Files", "Please upload files before merging.")
            return
        
        image_list = []
        pdf = FPDF()

        for file in self.files:
            ext = os.path.splitext(file)[1].lower()
            if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff"]:
                # Process images
                img = Image.open(file)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                image_list.append(img)
            elif ext == ".txt":
                # Convert text files to PDF
                self.convert_text_to_pdf(file, pdf)
            elif ext == ".docx":
                # Convert Word documents to PDF
                self.convert_docx_to_pdf(file, pdf)

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", 
                                                 filetypes=[("PDF file", "*.pdf")])
        if save_path:
            if image_list:
                first_image = image_list.pop(0)
                first_image.save(save_path, save_all=True, append_images=image_list)
            if pdf.page_no() > 0:
                pdf.output(save_path, "F")
            messagebox.showinfo("Success", f"PDF saved successfully as {os.path.basename(save_path)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FileToPDFApp(root)
    root.mainloop()
