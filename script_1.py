# Create an enhanced OCR GUI version with all the improvements
enhanced_ocr_gui = r'''
"""
Enhanced Electoral Roll OCR Converter - GUI Version with Advanced Debugging
==========================================================================

This enhanced GUI version includes:
- Multiple OCR processing strategies
- Raw OCR text debugging output
- Advanced image preprocessing
- Flexible voter record parsing
- Real-time progress tracking
- Comprehensive error handling

Requirements:
- pandas, openpyxl, pytesseract, Pillow, pdf2image, tkinter

Author: Generated AI Assistant
Date: 2025
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
import re
import os
import threading
import logging
import json
from datetime import datetime
from typing import List, Dict, Any, Optional

# OCR Libraries
try:
    import pytesseract
    from PIL import Image, ImageEnhance, ImageFilter
    import pdf2image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

class EnhancedElectoralRollGUI:
    """Enhanced GUI with advanced OCR processing and debugging"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Enhanced Electoral Roll OCR Converter - Debug Version")
        self.root.geometry("900x700")
        
        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Ready - Enhanced OCR with debugging enabled")
        self.dpi_var = tk.IntVar(value=300)
        self.debug_var = tk.BooleanVar(value=True)
        
        # Processing state
        self.is_processing = False
        self.current_page = 0
        self.total_pages = 0
        
        self.create_widgets()
        self.setup_logging()
        
    def create_widgets(self):
        """Create enhanced GUI widgets"""
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, 
                               text="Enhanced Electoral Roll OCR Converter", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))
        
        subtitle_label = ttk.Label(main_frame, 
                                  text="Advanced OCR processing with debugging for scanned PDFs",
                                  font=('Arial', 10), foreground="blue")
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 15))
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))
        
        # Input file
        ttk.Label(file_frame, text="Scanned PDF:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(file_frame, textvariable=self.input_file, width=70).grid(row=0, column=1, padx=5, pady=2)
        ttk.Button(file_frame, text="Browse...", command=self.select_input_file).grid(row=0, column=2, pady=2)
        
        # Output file
        ttk.Label(file_frame, text="Excel Output:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(file_frame, textvariable=self.output_file, width=70).grid(row=1, column=1, padx=5, pady=2)
        ttk.Button(file_frame, text="Browse...", command=self.select_output_file).grid(row=1, column=2, pady=2)
        
        # OCR Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="OCR Processing Settings", padding="10")
        settings_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        
        # DPI setting
        ttk.Label(settings_frame, text="Scan Quality (DPI):").grid(row=0, column=0, sticky=tk.W)
        dpi_scale = ttk.Scale(settings_frame, from_=150, to=600, variable=self.dpi_var, 
                             orient=tk.HORIZONTAL, length=200)
        dpi_scale.grid(row=0, column=1, padx=10)
        self.dpi_label = ttk.Label(settings_frame, textvariable=self.dpi_var)
        self.dpi_label.grid(row=0, column=2, padx=5)
        
        # Debug option
        debug_check = ttk.Checkbutton(settings_frame, text="Enable debug output (saves OCR text files)", 
                                     variable=self.debug_var)
        debug_check.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Quality info
        ttk.Label(settings_frame, text="Higher DPI = Better accuracy but slower processing", 
                 font=('Arial', 8), foreground="gray").grid(row=2, column=0, columnspan=3, pady=2)
        
        # Processing frame
        process_frame = ttk.Frame(main_frame)
        process_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Convert button
        self.convert_button = ttk.Button(process_frame, text="🔍 Start Enhanced OCR Processing", 
                                        command=self.start_conversion)
        self.convert_button.grid(row=0, column=0, padx=5)
        
        # Stop button
        self.stop_button = ttk.Button(process_frame, text="⏹ Stop Processing", 
                                     command=self.stop_conversion, state='disabled')
        self.stop_button.grid(row=0, column=1, padx=5)
        
        # Progress frame
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=5, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, length=600)
        self.progress_bar.grid(row=0, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))
        
        # Status labels
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var, font=('Arial', 9))
        self.status_label.grid(row=1, column=0, columnspan=3, pady=2)
        
        self.page_status = ttk.Label(progress_frame, text="", font=('Arial', 8), foreground="gray")
        self.page_status.grid(row=2, column=0, columnspan=3, pady=2)
        
        # Results notebook
        notebook = ttk.Notebook(main_frame)
        notebook.grid(row=6, column=0, columnspan=3, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Processing log tab
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="Processing Log")
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=100, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Debug info tab
        debug_frame = ttk.Frame(notebook)
        notebook.add(debug_frame, text="Debug Information")
        
        self.debug_text = scrolledtext.ScrolledText(debug_frame, height=15, width=100, wrap=tk.WORD)
        self.debug_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Statistics tab
        stats_frame = ttk.Frame(notebook)
        notebook.add(stats_frame, text="Extraction Statistics")
        
        self.stats_text = scrolledtext.ScrolledText(stats_frame, height=15, width=100, wrap=tk.WORD)
        self.stats_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5, pady=5)
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(6, weight=1)
        file_frame.columnconfigure(1, weight=1)
        settings_frame.columnconfigure(1, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        debug_frame.columnconfigure(0, weight=1)
        debug_frame.rowconfigure(0, weight=1)
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        
        # Check OCR setup
        self.check_ocr_setup()
        
        # Update DPI display
        self.dpi_var.trace('w', self.update_dpi_display)
        
    def update_dpi_display(self, *args):
        """Update DPI display label"""
        dpi_value = int(self.dpi_var.get())
        self.dpi_label.config(text=f"{dpi_value} DPI")
    
    def setup_logging(self):
        """Setup logging to capture in GUI"""
        self.gui_log_handler = GUILogHandler(self.log_text)
        logging.getLogger().addHandler(self.gui_log_handler)
        logging.getLogger().setLevel(logging.INFO)
    
    def check_ocr_setup(self):
        """Check OCR availability"""
        if not OCR_AVAILABLE:
            self.log_message("❌ OCR libraries not found!")
            self.log_message("Install: pip install pytesseract Pillow pdf2image")
            self.convert_button.config(state='disabled')
            return
        
        try:
            pytesseract.get_tesseract_version()
            self.log_message("✅ Enhanced OCR system ready!")
            self.log_message(f"Tesseract version: {pytesseract.get_tesseract_version()}")
            self.log_message("🔧 Debug mode enabled - OCR text will be saved for analysis")
        except Exception as e:
            self.log_message("❌ Tesseract OCR not found!")
            self.log_message("Install Tesseract OCR engine:")
            self.log_message("Windows: https://github.com/UB-Mannheim/tesseract/wiki")
            self.log_message("macOS: brew install tesseract")
            self.log_message("Linux: sudo apt-get install tesseract-ocr")
            self.convert_button.config(state='disabled')
    
    def select_input_file(self):
        """Select input PDF file"""
        filename = filedialog.askopenfilename(
            title="Select Scanned Electoral Roll PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            base_name = os.path.splitext(filename)[0]
            self.output_file.set(base_name + "_enhanced_ocr.xlsx")
    
    def select_output_file(self):
        """Select output Excel file"""
        filename = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.output_file.set(filename)
    
    def log_message(self, message):
        """Add message to log text area"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def debug_message(self, message):
        """Add message to debug text area"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.debug_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.debug_text.see(tk.END)
        self.root.update_idletasks()
    
    def stats_message(self, message):
        """Add message to statistics text area"""
        self.stats_text.insert(tk.END, f"{message}\n")
        self.stats_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_progress(self, value, status="", page_info=""):
        """Update progress display"""
        self.progress_var.set(value)
        if status:
            self.status_var.set(status)
        if page_info:
            self.page_status.config(text=page_info)
        self.root.update_idletasks()
    
    def start_conversion(self):
        """Start OCR conversion"""
        if not self.input_file.get() or not self.output_file.get():
            messagebox.showerror("Error", "Please select input PDF and output Excel files")
            return
        
        if not OCR_AVAILABLE:
            messagebox.showerror("Error", "OCR libraries not available")
            return
        
        # Clear previous results
        self.log_text.delete(1.0, tk.END)
        self.debug_text.delete(1.0, tk.END)
        self.stats_text.delete(1.0, tk.END)
        
        # Update UI
        self.is_processing = True
        self.convert_button.config(state='disabled')
        self.stop_button.config(state='normal')
        self.update_progress(0, "Initializing enhanced OCR processing...")
        
        # Start processing thread
        thread = threading.Thread(target=self.process_file)
        thread.daemon = True
        thread.start()
    
    def stop_conversion(self):
        """Stop OCR conversion"""
        self.is_processing = False
        self.update_progress(0, "Stopping processing...")
        self.log_message("🛑 Processing stopped by user")
    
    def process_file(self):
        """Main file processing method"""
        try:
            input_path = self.input_file.get()
            output_path = self.output_file.get()
            dpi = int(self.dpi_var.get())
            debug_enabled = self.debug_var.get()
            
            self.log_message("🔍 === ENHANCED OCR PROCESSING STARTED ===")
            self.log_message(f"📄 Input: {input_path}")
            self.log_message(f"📊 Output: {output_path}")
            self.log_message(f"🎯 Quality: {dpi} DPI")
            self.log_message(f"🔧 Debug mode: {'Enabled' if debug_enabled else 'Disabled'}")
            
            # Create extractor
            from enhanced_electoral_roll_ocr import EnhancedElectoralRollExtractor
            extractor = EnhancedElectoralRollExtractor(debug_mode=debug_enabled)
            
            # Convert PDF to images
            if not self.is_processing:
                return
                
            self.update_progress(5, "Converting PDF to images...", "Preparing for OCR processing")
            images = extractor.pdf_to_images(input_path, dpi=dpi)
            
            if not images:
                self.log_message("❌ Failed to convert PDF to images")
                return
            
            self.total_pages = len(images)
            self.log_message(f"📄 Successfully converted to {self.total_pages} images")
            
            # Process each page
            all_voters = []
            
            for page_num, image in enumerate(images, 1):
                if not self.is_processing:
                    self.log_message("🛑 Processing stopped")
                    return
                
                self.current_page = page_num
                progress = 10 + (70 * page_num / len(images))
                
                self.update_progress(
                    progress,
                    f"Processing page {page_num}/{len(images)} with OCR...",
                    f"Page {page_num}: Extracting text and analyzing voter records"
                )
                
                self.log_message(f"🔍 Processing page {page_num}/{len(images)}")
                
                try:
                    # Extract text with multiple strategies
                    page_text = extractor.extract_text_from_image_multi_config(image, page_num)
                    
                    if page_text.strip():
                        self.debug_message(f"Page {page_num}: Extracted {len(page_text)} characters")
                        
                        # Extract metadata from early pages
                        if page_num <= 3:
                            metadata = extractor.extract_metadata_flexible(page_text)
                            extractor.metadata.update(metadata)
                            if metadata:
                                self.debug_message(f"Page {page_num}: Found metadata: {metadata}")
                        
                        # Find voter records
                        voter_blocks = extractor.find_voter_records_flexible(page_text)
                        self.debug_message(f"Page {page_num}: Found {len(voter_blocks)} potential voter blocks")
                        
                        # Parse voter records
                        page_voters = []
                        for block_idx, block in enumerate(voter_blocks):
                            voter = extractor.parse_voter_record_flexible(
                                block, len(all_voters) + block_idx + 1, extractor.metadata
                            )
                            if voter:
                                page_voters.append(voter)
                        
                        all_voters.extend(page_voters)
                        
                        self.log_message(f"✅ Page {page_num}: Extracted {len(page_voters)} valid voter records")
                        
                    else:
                        self.log_message(f"⚠️ Page {page_num}: No text extracted")
                        
                except Exception as e:
                    self.log_message(f"❌ Page {page_num}: Processing error: {e}")
            
            if not self.is_processing:
                return
            
            # Save results
            self.update_progress(85, "Saving results to Excel...", "Formatting and saving voter data")
            
            if not all_voters:
                self.log_message("⚠️ WARNING: No voter records extracted!")
                self.log_message("Check debug files in ocr_debug_output/ folder")
                self.show_troubleshooting_tips()
                return
            
            # Save to Excel
            extractor.save_to_excel(all_voters, output_path)
            
            self.update_progress(100, "Enhanced OCR processing completed!", "")
            
            # Display results
            self.log_message("🎉 SUCCESS: Enhanced OCR processing completed!")
            self.log_message(f"📊 Total voters extracted: {len(all_voters)}")
            self.log_message(f"📁 Excel file saved: {output_path}")
            
            if debug_enabled:
                self.log_message("🔍 Debug files saved in ocr_debug_output/ folder")
            
            # Show statistics
            self.display_statistics(all_voters, extractor.metadata)
            
            # Ask to open file
            if messagebox.askyesno("Processing Complete",
                                 f"Enhanced OCR processing completed successfully!\n\n"
                                 f"Extracted {len(all_voters)} voter records from {len(images)} pages.\n\n"
                                 f"Would you like to open the Excel file?"):
                try:
                    os.startfile(output_path)
                except AttributeError:
                    try:
                        os.system(f"open '{output_path}'")
                    except:
                        os.system(f"xdg-open '{output_path}'")
            
        except Exception as e:
            self.log_message(f"❌ ERROR: {str(e)}")
            messagebox.showerror("Processing Error", f"An error occurred:\n\n{str(e)}")
        
        finally:
            # Reset UI
            self.is_processing = False
            self.convert_button.config(state='normal')
            self.stop_button.config(state='disabled')
    
    def display_statistics(self, voters, metadata):
        """Display extraction statistics"""
        self.stats_text.delete(1.0, tk.END)
        
        self.stats_message("📊 EXTRACTION STATISTICS")
        self.stats_message("=" * 50)
        
        # Basic stats
        self.stats_message(f"Total Records: {len(voters)}")
        self.stats_message(f"Processing Method: Enhanced OCR with multiple strategies")
        self.stats_message(f"Processing Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.stats_message("")
        
        # Metadata
        if metadata:
            self.stats_message("📋 CONSTITUENCY INFORMATION:")
            for key, value in metadata.items():
                self.stats_message(f"  {key.replace('_', ' ').title()}: {value}")
            self.stats_message("")
        
        # Gender distribution
        if voters:
            male_count = sum(1 for v in voters if v.get('gender') == 'M')
            female_count = sum(1 for v in voters if v.get('gender') == 'F')
            other_count = len(voters) - male_count - female_count
            
            self.stats_message("👥 GENDER DISTRIBUTION:")
            self.stats_message(f"  Male: {male_count} ({male_count/len(voters)*100:.1f}%)")
            self.stats_message(f"  Female: {female_count} ({female_count/len(voters)*100:.1f}%)")
            if other_count > 0:
                self.stats_message(f"  Other/Unknown: {other_count} ({other_count/len(voters)*100:.1f}%)")
            self.stats_message("")
            
            # Age distribution
            age_groups = {}
            for voter in voters:
                age_group = voter.get('age_group', 'Unknown')
                age_groups[age_group] = age_groups.get(age_group, 0) + 1
            
            self.stats_message("📊 AGE DISTRIBUTION:")
            for group, count in sorted(age_groups.items()):
                percentage = count / len(voters) * 100
                self.stats_message(f"  {group}: {count} ({percentage:.1f}%)")
            self.stats_message("")
            
            # Sample records
            self.stats_message("📝 SAMPLE EXTRACTED RECORDS:")
            for i, voter in enumerate(voters[:3], 1):
                self.stats_message(f"  {i}. {voter.get('name', 'N/A')} (Age: {voter.get('age', 'N/A')}, Gender: {voter.get('gender', 'N/A')})")
                self.stats_message(f"     EPIC: {voter.get('epic', 'N/A')}")
    
    def show_troubleshooting_tips(self):
        """Show troubleshooting information"""
        self.debug_message("🔧 TROUBLESHOOTING TIPS:")
        self.debug_message("1. Check raw OCR text files in ocr_debug_output/ folder")
        self.debug_message("2. Try increasing DPI setting to 400-600")
        self.debug_message("3. Verify PDF scan quality and contrast")
        self.debug_message("4. Check if pages are properly aligned (not rotated)")
        self.debug_message("5. Ensure PDF contains actual voter roll data")


class GUILogHandler(logging.Handler):
    """Custom logging handler for GUI display"""
    
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
    
    def emit(self, record):
        try:
            msg = self.format(record)
            def append():
                self.text_widget.insert(tk.END, msg + '\n')
                self.text_widget.see(tk.END)
            self.text_widget.after(0, append)
        except:
            pass


def main():
    """Main application entry point"""
    root = tk.Tk()
    
    # Check OCR availability
    if not OCR_AVAILABLE:
        messagebox.showerror("Missing Dependencies", 
                           "OCR libraries not found!\n\n"
                           "Please install:\n"
                           "pip install pytesseract Pillow pdf2image\n\n"
                           "Also install Tesseract OCR engine:\n"
                           "Windows: github.com/UB-Mannheim/tesseract/wiki\n"
                           "macOS: brew install tesseract\n"
                           "Linux: sudo apt-get install tesseract-ocr")
        root.destroy()
        return
    
    app = EnhancedElectoralRollGUI(root)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        root.destroy()


if __name__ == "__main__":
    main()
'''

# Save the enhanced OCR GUI
with open('enhanced_electoral_roll_ocr_gui.py', 'w', encoding='utf-8') as f:
    f.write(enhanced_ocr_gui)

print("✅ Enhanced OCR GUI created!")
print("- enhanced_electoral_roll_ocr_gui.py (Advanced OCR GUI with debugging)")
print()
print("📁 COMPLETE REWRITTEN SOLUTION:")
print("="*50)
print("🔍 enhanced_electoral_roll_ocr.py - Command-line version")
print("🖥️ enhanced_electoral_roll_ocr_gui.py - GUI version") 
print()
print("✨ NEW FEATURES IN REWRITTEN VERSION:")
print("• Saves raw OCR text to files for debugging")
print("• Multiple OCR processing strategies")
print("• Advanced image preprocessing")
print("• Flexible voter record detection")
print("• Better error correction and text cleaning")
print("• Comprehensive logging and progress tracking")
print("• Debug output folder with all processing details")
print()
print("🚀 TO USE:")
print("python enhanced_electoral_roll_ocr_gui.py")
print("OR")
print("python enhanced_electoral_roll_ocr.py input.pdf output.xlsx")