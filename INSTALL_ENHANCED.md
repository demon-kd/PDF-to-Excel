
# ENHANCED ELECTORAL ROLL OCR CONVERTER - INSTALLATION GUIDE

## Quick Setup (3 Steps)

### Step 1: Install Python Packages
```bash
pip install pandas openpyxl pytesseract Pillow pdf2image
```

### Step 2: Install Tesseract OCR Engine

**Windows:**
1. Download from: https://github.com/UB-Mannheim/tesseract/wiki
2. Install the .exe file
3. The program will auto-detect installation

**macOS:**
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install tesseract-ocr
```

### Step 3: Run the Enhanced OCR Converter

**Option A - GUI (Recommended):**
```bash
python enhanced_electoral_roll_ocr_gui.py
```

**Option B - Command Line:**
```bash
python enhanced_electoral_roll_ocr.py input.pdf output.xlsx
```

**Option C - Easy Launcher:**
```bash
python run_enhanced_ocr.py
```

## What the Enhanced Version Does

### 🔍 Advanced OCR Processing
- Uses multiple OCR strategies for better accuracy
- Advanced image preprocessing (contrast, sharpness, scaling)
- Saves raw OCR text to files for debugging
- Better handles poor quality scans

### 📊 Flexible Data Extraction  
- Multiple voter record detection methods
- Improved regex patterns for various PDF layouts
- Advanced text cleaning and error correction
- Handles missing or corrupted data gracefully

### 🐛 Comprehensive Debugging
- Creates `ocr_debug_output/` folder with:
  - Original page images (`page_001_original.png`)
  - Processed images (`page_001_processed.png`) 
  - Raw OCR text files (`page_001_raw_ocr.txt`)
  - Combined text file (`all_pages_combined_text.txt`)
  - Extraction summary (`extraction_summary.json`)

### 📈 Better Results
- Higher accuracy voter record extraction
- Improved handling of OCR errors
- Better metadata detection
- More comprehensive error reporting

## If No Records Are Found

1. **Check debug files** in `ocr_debug_output/` folder
2. **Review raw OCR text** files to see what was extracted
3. **Increase DPI** setting to 400-600 for better quality
4. **Verify PDF quality** - ensure good contrast and alignment
5. **Share OCR text files** for custom parsing if needed

## Troubleshooting

### "Tesseract not found" Error
- Follow Step 2 installation instructions above
- Check that Tesseract is in your system PATH

### "No voter records extracted" Warning
- Check the `page_XXX_raw_ocr.txt` files to see actual OCR output
- Try higher DPI setting (400-600)
- Ensure PDF is actually scanned (not text-based)

### Poor OCR Quality
- Use higher quality source scans (300+ DPI)
- Ensure good contrast between text and background
- Check that pages are not rotated or skewed

## Files in Enhanced Solution

- `enhanced_electoral_roll_ocr.py` - Main command-line program
- `enhanced_electoral_roll_ocr_gui.py` - GUI version with debugging
- `run_enhanced_ocr.py` - Easy launcher with requirement checking
- `INSTALL_ENHANCED.md` - This installation guide

The enhanced version is specifically designed to handle challenging scanned PDFs and provide detailed debugging information to help resolve extraction issues.
