import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from fpdf import FPDF
import os
from docx import Document
import customtkinter as ctk  # You'll need to install this: pip install customtkinter

# Set appearance mode and default color theme
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class MaterialPDFConverter(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("Material PDF Converter")
        self.geometry("700x500")
        self.minsize(600, 500)
        
        # Set the overall theme
        self._set_appearance_mode("light")
        
        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)  # Title doesn't expand
        self.grid_rowconfigure(1, weight=1)  # File list expands
        self.grid_rowconfigure(2, weight=0)  # Buttons don't expand
        
        # Create fonts and colors
        self.title_font = ctk.CTkFont(family="Roboto", size=24, weight="bold")
        self.button_font = ctk.CTkFont(family="Roboto", size=14)
        self.list_font = ctk.CTkFont(family="Roboto", size=12)
        
        # Colors from the provided Material Design 3 color scheme
        # Primary color family (green)
        self.primary_color = "#3E713E"  # P-40 (Primary)
        self.primary_on_color = "#FFFFFF"  # Text on primary
        self.primary_container = "#B5E5A9"  # P-90 (Primary Container)
        self.primary_on_container = "#0A2F0A"  # P-10 (On Primary Container)
        
        # Secondary color family (olive)
        self.secondary_color = "#5A6349"  # S-40 (Secondary)
        self.secondary_on_color = "#FFFFFF"  # Text on secondary
        self.secondary_container = "#D7E6B9"  # S-90 (Secondary Container)
        self.secondary_on_container = "#171E0B"  # S-10 (On Secondary Container)
        
        # Tertiary color family (teal)
        self.tertiary_color = "#396569"  # T-40 (Tertiary)
        self.tertiary_on_color = "#FFFFFF"  # Text on tertiary
        self.tertiary_container = "#B8E8ED"  # T-90 (Tertiary Container)
        self.tertiary_on_container = "#0A2628"  # T-10 (On Tertiary Container)
        
        # Error color family (red)
        self.error_color = "#BA1A1A"  # E-40 (Error)
        self.error_container = "#FFDAD6"  # E-90 (Error Container)
        
        # Neutral color family
        self.surface_color = "#F5F5F0"  # N-98 (Surface)
        self.surface_dim = "#E3E3DE"  # N-87 (Surface Dim)
        self.surface_container = "#EBEBEB"  # N-94 (Surface Container)
        self.on_surface = "#1A1C18"  # N-10 (On Surface)
        
        # Create header
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=self.primary_color)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 10))
        self.header_frame.grid_columnconfigure(0, weight=1)
        
        self.title_label = ctk.CTkLabel(
            self.header_frame, 
            text="PDF Converter", 
            font=self.title_font,
            text_color=self.primary_on_color
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=15)
        
        # Create main content frame
        self.content_frame = ctk.CTkFrame(self, fg_color=self.surface_color)
        self.content_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        # Create file list frame with scroll
        self.list_frame = ctk.CTkFrame(self.content_frame, fg_color=self.surface_container)
        self.list_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.list_frame.grid_columnconfigure(0, weight=1)
        self.list_frame.grid_rowconfigure(0, weight=1)
        
        # Create custom listbox (CTk doesn't have one, so we'll style a normal one)
        self.file_listbox_frame = ctk.CTkFrame(self.list_frame, fg_color=self.surface_container)
        self.file_listbox_frame.grid(row=0, column=0, sticky="nsew")
        self.file_listbox_frame.grid_rowconfigure(0, weight=1)
        self.file_listbox_frame.grid_columnconfigure(0, weight=1)
        
        # File list with scrollbar
        self.file_listbox = tk.Listbox(
            self.file_listbox_frame,
            selectbackground=self.primary_color,
            selectforeground=self.primary_on_color,
            font=self.list_font,
            borderwidth=0,
            highlightthickness=0,
            bg=self.surface_color,
            fg=self.on_surface
        )
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbar for listbox
        self.scrollbar = ttk.Scrollbar(self.file_listbox_frame, orient="vertical", command=self.file_listbox.yview)
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=self.scrollbar.set)
        
        # List label
        self.list_label = ctk.CTkLabel(
            self.list_frame, 
            text="Files to Convert", 
            font=ctk.CTkFont(family="Roboto", size=16, weight="bold"),
            text_color=self.on_surface,
            anchor="w"
        )
        self.list_label.grid(row=0, column=0, padx=10, pady=(0, 10), sticky="nw")
        
        # Create button frame
        self.button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.button_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.button_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Create buttons
        self.add_button = ctk.CTkButton(
            self.button_frame,
            text="Add Files",
            font=self.button_font,
            command=self.add_file,
            fg_color=self.primary_color,
            text_color=self.primary_on_color,
            hover_color=self.primary_color,  # Slightly darker shade on hover
            corner_radius=24,
            height=48
        )
        self.add_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        self.remove_button = ctk.CTkButton(
            self.button_frame,
            text="Remove Selected",
            font=self.button_font,
            command=self.remove_file,
            fg_color=self.secondary_color,
            text_color=self.secondary_on_color,
            hover_color=self.secondary_color,  # Slightly darker shade on hover
            state="disabled",
            corner_radius=24,
            height=48
        )
        self.remove_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        self.merge_button = ctk.CTkButton(
            self.button_frame,
            text="Convert to PDF",
            font=self.button_font,
            command=self.merge_to_pdf,
            fg_color=self.tertiary_color,  # Using tertiary color for action
            text_color=self.tertiary_on_color,
            hover_color=self.tertiary_color,  # Slightly darker shade on hover
            state="disabled",
            corner_radius=24,
            height=48
        )
        self.merge_button.grid(row=0, column=2, padx=10, pady=10, sticky="ew")
        
        # Status bar at the bottom
        self.status_var = tk.StringVar(value="Ready to convert files")
        self.status_bar = ctk.CTkLabel(
            self, 
            textvariable=self.status_var,
            font=ctk.CTkFont(family="Roboto", size=12),
            anchor="w"
        )
        self.status_bar.grid(row=3, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        # File list to store the selected files
        self.files = []
        
        # Bind events
        self.file_listbox.bind("<<ListboxSelect>>", self.on_select)
        
        # Add file type icons (if you want to add this feature later)
        self.file_icons = {
            ".jpg": "🖼️",
            ".jpeg": "🖼️",
            ".png": "🖼️",
            ".bmp": "🖼️",
            ".tiff": "🖼️",
            ".txt": "📄",
            ".docx": "📝",
            ".pdf": "📑"
        }

    def on_select(self, event=None):
        if self.file_listbox.curselection():
            self.remove_button.configure(state="normal")
        else:
            self.remove_button.configure(state="disabled")

    def add_file(self):
        files = filedialog.askopenfilenames(
            title="Select Files to Convert", 
            filetypes=[
                ("Supported Files", "*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.txt;*.docx;*.pdf"),
                ("Image Files", "*.jpg;*.jpeg;*.png;*.bmp;*.tiff"),
                ("Text Files", "*.txt"),
                ("Word Documents", "*.docx"),
                ("PDF Files", "*.pdf")
            ]
        )
        
        if files:
            if len(self.files) + len(files) > 30:
                messagebox.showwarning("Too many files", "You can only upload up to 30 files.")
                # Add as many as we can
                remaining = 30 - len(self.files)
                files = files[:remaining]
            
            # Add the files to our list and listbox
            for file in files:
                self.files.append(file)
                ext = os.path.splitext(file)[1].lower()
                icon = self.file_icons.get(ext, "📄")
                self.file_listbox.insert(tk.END, f"{icon} {os.path.basename(file)}")
            
            self.merge_button.configure(state="normal")
            self.status_var.set(f"{len(self.files)} files ready to convert")

    def remove_file(self):
        selected_indices = self.file_listbox.curselection()
        if selected_indices:
            # Delete in reverse order to avoid index shifting
            for index in sorted(selected_indices, reverse=True):
                self.file_listbox.delete(index)
                del self.files[index]
            
            # Update button states
            if not self.files:
                self.merge_button.configure(state="disabled")
                self.remove_button.configure(state="disabled")
                self.status_var.set("Ready to convert files")
            else:
                self.status_var.set(f"{len(self.files)} files ready to convert")

    def convert_text_to_pdf(self, text_file, pdf):
        try:
            with open(text_file, 'r', encoding='utf-8') as file:
                text_content = file.read()
                pdf.add_page()
                pdf.set_auto_page_break(auto=True, margin=15)
                pdf.set_font("Arial", size=12)
                pdf.multi_cell(0, 10, text_content)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert text file: {str(e)}")
            return False

    def convert_docx_to_pdf(self, docx_file, pdf):
        try:
            doc = Document(docx_file)
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.set_font("Arial", size=12)
            
            for paragraph in doc.paragraphs:
                pdf.multi_cell(0, 10, paragraph.text)
                pdf.ln(5)  # Line break between paragraphs
            
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert DOCX file: {str(e)}")
            return False

    def merge_to_pdf(self):
        if not self.files:
            messagebox.showerror("No Files", "Please add files before converting to PDF.")
            return
        
        # Ask for save location
        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            filetypes=[("PDF file", "*.pdf")],
            title="Save PDF As"
        )
        
        if not save_path:
            return  # User cancelled
        
        # Show progress
        self.status_var.set("Converting files to PDF...")
        self.update_idletasks()
        
        # Process files based on type
        image_files = []
        pdf = FPDF()
        
        try:
            for file in self.files:
                ext = os.path.splitext(file)[1].lower()
                
                if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff"]:
                    # Collect image files for PIL processing
                    img = Image.open(file)
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    image_files.append(img)
                
                elif ext == ".txt":
                    # Process text files
                    self.convert_text_to_pdf(file, pdf)
                
                elif ext == ".docx":
                    # Process Word documents
                    self.convert_docx_to_pdf(file, pdf)
                
                elif ext == ".pdf":
                    # We'd need PyPDF2 or similar to merge existing PDFs
                    # This can be added as a future enhancement
                    messagebox.showinfo("Info", "PDF merging requires additional libraries. Skipping PDF files for now.")
            
            # Save the results
            pdf_created = False
            
            if image_files:
                # Save the images as a PDF
                first_img = image_files.pop(0)
                first_img.save(save_path, save_all=True, append_images=image_files)
                pdf_created = True
            
            if pdf.page_no() > 0:
                if pdf_created:
                    # We need to handle the case where both images and text/docx files were processed
                    # This would require PyPDF2 to merge the two PDFs
                    temp_path = save_path + ".temp.pdf"
                    pdf.output(temp_path)
                    messagebox.showinfo("Partial Success", 
                                       "PDF saved with images only. Text and DOCX files require additional libraries to merge.")
                else:
                    pdf.output(save_path)
                    pdf_created = True
            
            if pdf_created:
                self.status_var.set(f"PDF saved successfully as {os.path.basename(save_path)}")
                # Show success message
                messagebox.showinfo("Success", f"PDF created successfully:\n{save_path}")
            else:
                self.status_var.set("No files were converted")
                messagebox.showerror("Error", "No files could be converted to PDF")
                
        except Exception as e:
            self.status_var.set("Error during conversion")
            messagebox.showerror("Error", f"Failed to create PDF: {str(e)}")


if __name__ == "__main__":
    app = MaterialPDFConverter()
    app.mainloop()
