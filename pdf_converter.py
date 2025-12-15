"""
PDF Converter - Convert Office files to PDF using LibreOffice
Author: Leo Lynn (https://github.com/leolynn7)
Version: 1.3
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess
import threading
from pathlib import Path
import webbrowser

class LibreOfficePDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Converter (LibreOffice)")
        self.root.geometry("600x600")
        
        # Dark theme colors
        bg_color = '#2b2b2b'
        fg_color = '#ffffff'
        btn_color = '#007acc'
        accent_color = '#4ec9b0'
        
        self.root.configure(bg=bg_color)
        
        # Variables
        self.files_to_convert = []
        self.output_dir = tk.StringVar(value=os.path.expanduser("~/Desktop"))
        
        # Create GUI with proper layout management
        self.create_widgets(bg_color, fg_color, btn_color, accent_color)
        
        # Check for LibreOffice
        self.check_libreoffice()
        
        # Configure grid weights for proper stretching
        self.configure_grid()
    
    def configure_grid(self):
        """Configure grid weights for proper stretching"""
        # Make the main frame expandable
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
    
    def create_widgets(self, bg_color, fg_color, btn_color, accent_color):
        """Create the GUI interface with proper layout"""
        
        # Main container frame using grid
        main_frame = tk.Frame(self.root, bg=bg_color)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        # Configure main_frame grid weights
        main_frame.grid_rowconfigure(0, weight=0)  # Title
        main_frame.grid_rowconfigure(1, weight=0)  # Instructions
        main_frame.grid_rowconfigure(2, weight=0)  # Output folder section
        main_frame.grid_rowconfigure(3, weight=0)  # Files label
        main_frame.grid_rowconfigure(4, weight=1)  # File list (expands)
        main_frame.grid_rowconfigure(5, weight=0)  # File buttons
        main_frame.grid_rowconfigure(6, weight=0)  # Convert button
        main_frame.grid_rowconfigure(7, weight=0)  # Status
        main_frame.grid_rowconfigure(8, weight=0)  # Spacer
        main_frame.grid_rowconfigure(9, weight=0)  # Credit section
        main_frame.grid_columnconfigure(0, weight=1)
        
        row = 0
        
        # Title
        title = tk.Label(main_frame, 
                        text="📄 PDF Converter", 
                        font=('Arial', 20, 'bold'),
                        bg=bg_color, 
                        fg=fg_color)
        title.grid(row=row, column=0, pady=(0, 10), sticky="w")
        row += 1
        
        # Instructions
        instr = tk.Label(main_frame,
                        text="Uses LibreOffice for perfect PDF conversions\n"
                             "Preserves all formatting: bullets, tables, images",
                        bg=bg_color,
                        fg='#cccccc',
                        font=('Arial', 10))
        instr.grid(row=row, column=0, pady=(0, 20), sticky="w")
        row += 1
        
        # Output Folder Section
        output_frame = tk.Frame(main_frame, bg=bg_color)
        output_frame.grid(row=row, column=0, pady=(0, 15), sticky="ew")
        output_frame.grid_columnconfigure(0, weight=1)
        row += 1
        
        tk.Label(output_frame, 
                text="Output Folder:", 
                bg=bg_color, 
                fg=fg_color,
                font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky="w")
        
        # Output folder path with browse button
        path_frame = tk.Frame(output_frame, bg=bg_color)
        path_frame.grid(row=1, column=0, pady=(5, 0), sticky="ew")
        path_frame.grid_columnconfigure(0, weight=1)
        
        # Current path display
        self.path_display = tk.Label(path_frame,
                                    textvariable=self.output_dir,
                                    bg='#3c3c3c',
                                    fg='#cccccc',
                                    font=('Monospace', 9),
                                    padx=10,
                                    pady=8,
                                    anchor=tk.W,
                                    relief=tk.FLAT)
        self.path_display.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        # Browse button for output folder
        browse_btn = tk.Button(path_frame,
                              text="📁 Browse",
                              command=self.select_output_folder,
                              bg='#555555',
                              fg=fg_color,
                              activebackground='#666666',
                              activeforeground=fg_color,
                              relief=tk.FLAT,
                              padx=15,
                              pady=5,
                              cursor='hand2',
                              font=('Arial', 9))
        browse_btn.grid(row=0, column=1, sticky="e")
        
        # Files to Convert Label
        tk.Label(main_frame, 
                text="Files to Convert:", 
                bg=bg_color, 
                fg=fg_color,
                font=('Arial', 10, 'bold')).grid(row=row, column=0, pady=(0, 5), sticky="w")
        row += 1
        
        # File listbox with scrollbar in a frame
        list_container = tk.Frame(main_frame, bg=bg_color)
        list_container.grid(row=row, column=0, pady=(0, 10), sticky="nsew")
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1)
        
        # Scrollbar for listbox
        scrollbar = tk.Scrollbar(list_container)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # File listbox
        self.file_listbox = tk.Listbox(list_container,
                                      bg='#3c3c3c',
                                      fg=fg_color,
                                      selectbackground=btn_color,
                                      selectforeground=fg_color,
                                      font=('Monospace', 10),
                                      relief=tk.FLAT,
                                      borderwidth=2,
                                      yscrollcommand=scrollbar.set)
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.config(command=self.file_listbox.yview)
        row += 1
        
        # File Management Buttons - CENTERED
        file_btn_frame = tk.Frame(main_frame, bg=bg_color)
        file_btn_frame.grid(row=row, column=0, pady=(0, 15), sticky="")
        row += 1
        
        # Center the buttons in the frame
        file_btn_frame.grid_columnconfigure(0, weight=1)
        file_btn_frame.grid_columnconfigure(4, weight=1)
        
        # File operation buttons with centered layout
        self.add_btn = self.create_button(file_btn_frame, "📁 Add Files", self.add_files, btn_color, fg_color)
        self.add_btn.grid(row=0, column=1, padx=(0, 10))
        
        self.remove_btn = self.create_button(file_btn_frame, "🗑️ Remove Selected", self.remove_files, '#d9534f', fg_color)
        self.remove_btn.grid(row=0, column=2, padx=(0, 10))
        
        self.clear_btn = self.create_button(file_btn_frame, "🧹 Clear All", self.clear_files, '#5bc0de', fg_color)
        self.clear_btn.grid(row=0, column=3)
        
        # Convert Button - CENTERED
        convert_frame = tk.Frame(main_frame, bg=bg_color)
        convert_frame.grid(row=row, column=0, pady=(0, 15), sticky="ew")
        convert_frame.grid_columnconfigure(0, weight=1)
        convert_frame.grid_columnconfigure(2, weight=1)
        row += 1
        
        self.convert_btn = tk.Button(convert_frame,
                                    text="⚡ CONVERT TO PDF",
                                    command=self.convert_files,
                                    bg='#5cb85c',
                                    fg=fg_color,
                                    activebackground='#4cae4c',
                                    activeforeground=fg_color,
                                    relief=tk.FLAT,
                                    padx=30,
                                    pady=15,
                                    cursor='hand2',
                                    font=('Arial', 13, 'bold'))
        self.convert_btn.grid(row=0, column=1, sticky="")
        
        # Status label
        self.status_label = tk.Label(main_frame,
                                    text="Ready to convert files",
                                    bg=bg_color,
                                    fg='#aaaaaa',
                                    font=('Arial', 10))
        self.status_label.grid(row=row, column=0, pady=(0, 20), sticky="w")
        row += 1
        
        # Spacer to push credit to bottom
        spacer = tk.Frame(main_frame, bg=bg_color, height=20)
        spacer.grid(row=row, column=0, sticky="ew")
        spacer.grid_rowconfigure(0, weight=1)
        row += 1
        
        # CREDIT SECTION - Now properly positioned at bottom
        credit_frame = tk.Frame(main_frame, bg=bg_color)
        credit_frame.grid(row=row, column=0, sticky="ew", pady=(10, 0))
        row += 1
        
        # Add a subtle separator line
        separator = tk.Frame(credit_frame, height=1, bg='#444444')
        separator.pack(fill=tk.X, pady=(0, 10))
        
        # Credit container - CENTERED
        credit_container = tk.Frame(credit_frame, bg=bg_color)
        credit_container.pack()
        
        # Created by text
        created_by = tk.Label(credit_container,
                             text="Created by ",
                             bg=bg_color,
                             fg='#888888',
                             font=('Arial', 10))
        created_by.pack(side=tk.LEFT)
        
        # Clickable name with link - LEO LYNN
        self.leo_link = tk.Label(credit_container,
                                text="Leo Lynn",
                                bg=bg_color,
                                fg=accent_color,
                                font=('Arial', 10, 'underline', 'bold'),
                                cursor="hand2")
        self.leo_link.pack(side=tk.LEFT)
        
        # Bind click event to open GitHub
        self.leo_link.bind("<Button-1>", lambda e: self.open_github())
        self.leo_link.bind("<Enter>", lambda e: self.leo_link.config(fg='#7ec1ff'))
        self.leo_link.bind("<Leave>", lambda e: self.leo_link.config(fg=accent_color))
        
        # GitHub link text
        github_text = tk.Label(credit_container,
                              text=" (",
                              bg=bg_color,
                              fg='#888888',
                              font=('Arial', 10))
        github_text.pack(side=tk.LEFT)
        
        # GitHub link
        github_link = tk.Label(credit_container,
                              text="GitHub",
                              bg=bg_color,
                              fg='#4ec9b0',
                              font=('Arial', 10, 'underline'),
                              cursor="hand2")
        github_link.pack(side=tk.LEFT)
        github_link.bind("<Button-1>", lambda e: self.open_github())
        github_link.bind("<Enter>", lambda e: github_link.config(fg='#7ec1ff'))
        github_link.bind("<Leave>", lambda e: github_link.config(fg='#4ec9b0'))
        
        # Closing parenthesis
        closing_paren = tk.Label(credit_container,
                                text=")",
                                bg=bg_color,
                                fg='#888888',
                                font=('Arial', 10))
        closing_paren.pack(side=tk.LEFT)
        
        # Version label
        version_label = tk.Label(credit_container,
                                text=" • v1.3",
                                bg=bg_color,
                                fg='#666666',
                                font=('Arial', 9))
        version_label.pack(side=tk.LEFT, padx=(5, 0))
        
        # GitHub octocat icon
        octocat = tk.Label(credit_container,
                          text=" 🐙",
                          bg=bg_color,
                          fg=accent_color,
                          font=('Arial', 12),
                          cursor="hand2")
        octocat.pack(side=tk.LEFT)
        octocat.bind("<Button-1>", lambda e: self.open_github())
    
    def create_button(self, parent, text, command, bg, fg):
        """Create a styled button"""
        btn = tk.Button(parent,
                       text=text,
                       command=command,
                       bg=bg,
                       fg=fg,
                       activebackground=bg,
                       activeforeground=fg,
                       relief=tk.FLAT,
                       padx=15,
                       pady=8,
                       cursor='hand2',
                       font=('Arial', 10))
        return btn
    
    def select_output_folder(self):
        """Select output folder using file dialog"""
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_dir.get()
        )
        if folder:  # Only update if user didn't cancel
            self.output_dir.set(folder)
    
    def open_github(self):
        """Open GitHub profile in web browser"""
        webbrowser.open("https://github.com/leolynn7")
    
    def check_libreoffice(self):
        """Check if LibreOffice is installed"""
        try:
            result = subprocess.run(['which', 'libreoffice'], 
                                   capture_output=True, 
                                   text=True)
            if result.returncode != 0:
                self.show_install_instructions()
                return False
            return True
        except:
            self.show_install_instructions()
            return False
    
    def show_install_instructions(self):
        """Show LibreOffice installation instructions"""
        messagebox.showinfo("LibreOffice Required",
                          "LibreOffice is required for perfect PDF conversions.\n\n"
                          "Install it with:\n\n"
                          "Ubuntu/Debian: sudo apt install libreoffice\n"
                          "Fedora: sudo dnf install libreoffice\n"
                          "Arch: sudo pacman -S libreoffice-fresh\n\n"
                          "Or download from: https://libreoffice.org")
        
        # Disable convert button
        self.convert_btn.config(state='disabled', bg='#666666')
    
    def add_files(self):
        """Add files to convert"""
        filetypes = [
            ("Office Files", "*.docx *.doc *.xlsx *.xls *.odt *.ods *.ppt *.pptx"),
            ("All Files", "*.*")
        ]
        
        files = filedialog.askopenfilenames(filetypes=filetypes)
        
        for file in files:
            if file not in self.files_to_convert:
                self.files_to_convert.append(file)
                display_name = os.path.basename(file)
                # Truncate long filenames
                if len(display_name) > 50:
                    display_name = display_name[:47] + "..."
                self.file_listbox.insert(tk.END, f"📄 {display_name}")
    
    def remove_files(self):
        """Remove selected files"""
        selection = self.file_listbox.curselection()
        for index in reversed(selection):
            self.files_to_convert.pop(index)
            self.file_listbox.delete(index)
    
    def clear_files(self):
        """Clear all files"""
        self.files_to_convert.clear()
        self.file_listbox.delete(0, tk.END)
    
    def convert_files(self):
        """Convert files using LibreOffice"""
        if not self.files_to_convert:
            messagebox.showwarning("No Files", "Please add files first.")
            return
        
        # Check if output directory exists
        output_dir = self.output_dir.get()
        if not os.path.exists(output_dir):
            create = messagebox.askyesno("Create Folder", 
                                        f"Output folder doesn't exist:\n{output_dir}\n\nCreate it?")
            if create:
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except Exception as e:
                    messagebox.showerror("Error", f"Cannot create folder:\n{str(e)}")
                    return
            else:
                return
        
        # Start conversion in thread
        thread = threading.Thread(target=lambda: self.run_conversion(output_dir), daemon=True)
        thread.start()
    
    def run_conversion(self, output_dir):
        """Run the actual conversion"""
        total = len(self.files_to_convert)
        successful = 0
        
        # Update UI to show conversion starting
        self.root.after(0, self.convert_btn.config, {'state': 'disabled', 'bg': '#666666'})
        self.root.after(0, self.status_label.config, {'text': f"Starting conversion of {total} files..."})
        
        for i, file in enumerate(self.files_to_convert):
            # Update status
            filename = os.path.basename(file)
            self.root.after(0, self.update_status, f"Converting {i+1}/{total}: {filename}")
            
            try:
                # Use LibreOffice to convert
                cmd = [
                    'libreoffice',
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', output_dir,
                    file
                ]
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                
                if result.returncode == 0:
                    successful += 1
                    # Update listbox to show success
                    self.root.after(0, self.mark_as_converted, i)
                else:
                    error_msg = result.stderr[:100] if result.stderr else "Unknown error"
                    self.root.after(0, self.show_error, 
                                  f"Failed to convert: {filename}\nError: {error_msg}")
            
            except subprocess.TimeoutExpired:
                self.root.after(0, self.show_error, f"Timeout: {filename} took too long")
            except Exception as e:
                self.root.after(0, self.show_error, f"Error converting {filename}: {str(e)[:50]}")
        
        # Complete
        self.root.after(0, self.conversion_complete, successful, total, output_dir)
    
    def mark_as_converted(self, index):
        """Mark a file as converted in the listbox"""
        current_text = self.file_listbox.get(index)
        if not current_text.startswith("✅ "):
            # Replace the file icon with checkmark
            new_text = "✅ " + current_text[2:]
            self.file_listbox.delete(index)
            self.file_listbox.insert(index, new_text)
    
    def update_status(self, message):
        """Update status label"""
        self.status_label.config(text=message)
    
    def show_error(self, message):
        """Show error message"""
        messagebox.showerror("Error", message)
    
    def conversion_complete(self, successful, total, output_dir):
        """Handle completion"""
        # Re-enable convert button
        self.root.after(0, self.convert_btn.config, {'state': 'normal', 'bg': '#5cb85c'})
        
        if successful == total:
            message = f"✅ Successfully converted {successful} files!"
            details = f"All {successful} files converted successfully!\n\nSaved to:\n{output_dir}"
        elif successful > 0:
            message = f"⚠️ Converted {successful} out of {total} files"
            details = f"Converted {successful}/{total} files\n\nSaved to:\n{output_dir}\n\nCheck error messages for failed files."
        else:
            message = "❌ Conversion failed!"
            details = f"Failed to convert all {total} files.\n\nCheck error messages above."
        
        self.status_label.config(text=message)
        messagebox.showinfo("Conversion Complete", details)

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    
    # Set window icon if available
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
    
    app = LibreOfficePDFConverter(root)
    
    # Center window
    root.update_idletasks()
    width = 600
    height = 600
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    # Make window resizable with minimum size
    root.resizable(True, True)
    root.minsize(550, 550)
    
    # Configure root window to expand properly
    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(0, weight=1)
    
    root.mainloop()