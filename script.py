# Create an improved OCR program with better debugging and parsing capabilities
improved_ocr_program = r'''
"""
Enhanced Electoral Roll Scanned PDF to Excel Converter with Advanced OCR
========================================================================

This improved version includes:
- Raw OCR text debugging output
- More flexible voter record parsing
- Better error handling and text cleaning
- Multiple parsing strategies
- Detailed logging and progress tracking

Requirements:
- pandas, openpyxl, pytesseract, Pillow, pdf2image
- Tesseract OCR engine

Usage:
    python enhanced_electoral_roll_ocr.py input_file.pdf output_file.xlsx

Author: Generated AI Assistant
Date: 2025
"""

import pandas as pd
import re
import sys
import os
import logging
from typing import List, Dict, Any, Optional
import json
from datetime import datetime

# OCR and Image Processing Libraries
try:
    import pytesseract
    from PIL import Image, ImageEnhance, ImageFilter
    import pdf2image
    OCR_AVAILABLE = True
except ImportError as e:
    print(f"❌ OCR libraries not available: {e}")
    print("Please install: pip install pytesseract Pillow pdf2image")
    OCR_AVAILABLE = False

class EnhancedElectoralRollExtractor:
    """
    Enhanced extractor with improved OCR processing and debugging capabilities
    """
    
    def __init__(self, debug_mode=True):
        self.data = []
        self.metadata = {}
        self.debug_mode = debug_mode
        self.output_dir = "ocr_debug_output"
        self.setup_logging()
        self.setup_tesseract()
        self.create_debug_directory()
        
    def setup_logging(self):
        """Setup comprehensive logging"""
        log_format = '%(asctime)s - %(levelname)s - %(message)s'
        logging.basicConfig(
            level=logging.INFO,
            format=log_format,
            handlers=[
                logging.FileHandler('electoral_ocr.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def create_debug_directory(self):
        """Create directory for debug outputs"""
        if self.debug_mode and not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
            self.logger.info(f"Created debug directory: {self.output_dir}")
    
    def setup_tesseract(self):
        """Setup Tesseract with multiple path attempts"""
        tesseract_paths = [
            r'C:\Program Files\Tesseract-OCR\tesseract.exe',
            r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
            '/usr/bin/tesseract',
            '/usr/local/bin/tesseract',
            '/opt/homebrew/bin/tesseract',
            'tesseract'  # Try system PATH
        ]
        
        for path in tesseract_paths:
            try:
                if path == 'tesseract':
                    # Test system PATH
                    pytesseract.get_tesseract_version()
                    self.logger.info("Using Tesseract from system PATH")
                    return
                elif os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    pytesseract.get_tesseract_version()
                    self.logger.info(f"Found Tesseract at: {path}")
                    return
            except Exception:
                continue
        
        self.logger.error("❌ Tesseract OCR not found!")
        print("\n🔧 INSTALL TESSERACT OCR:")
        print("Windows: https://github.com/UB-Mannheim/tesseract/wiki")
        print("macOS: brew install tesseract")
        print("Linux: sudo apt-get install tesseract-ocr")
    
    def pdf_to_images(self, pdf_path: str, dpi: int = 300) -> List[Image.Image]:
        """Convert PDF to images with error handling"""
        self.logger.info(f"Converting PDF to images (DPI: {dpi})...")
        
        try:
            images = pdf2image.convert_from_path(
                pdf_path,
                dpi=dpi,
                fmt='PNG',
                thread_count=4,
                first_page=None,
                last_page=None
            )
            
            self.logger.info(f"✅ Converted {len(images)} pages to images")
            
            # Save images for debugging if enabled
            if self.debug_mode:
                for i, img in enumerate(images, 1):
                    img_path = os.path.join(self.output_dir, f"page_{i:03d}_original.png")
                    img.save(img_path)
                    self.logger.info(f"Saved debug image: {img_path}")
            
            return images
            
        except Exception as e:
            self.logger.error(f"❌ PDF conversion failed: {e}")
            raise Exception(f"Failed to convert PDF to images: {e}")
    
    def preprocess_image_advanced(self, image: Image.Image, page_num: int) -> Image.Image:
        """Advanced image preprocessing for better OCR"""
        
        self.logger.info(f"Preprocessing image for page {page_num}")
        
        # Convert to grayscale
        if image.mode != 'L':
            image = image.convert('L')
        
        # Enhance contrast
        enhancer = ImageEnhance.Contrast(image)
        image = enhancer.enhance(1.5)  # Increase contrast
        
        # Enhance sharpness
        enhancer = ImageEnhance.Sharpness(image)
        image = enhancer.enhance(2.0)  # Increase sharpness
        
        # Apply slight blur to reduce noise
        image = image.filter(ImageFilter.GaussianBlur(radius=0.5))
        
        # Resize if too small
        width, height = image.size
        if width < 1500:
            scale_factor = 1500 / width
            new_width = int(width * scale_factor)
            new_height = int(height * scale_factor)
            image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
            self.logger.info(f"Resized image to {new_width}x{new_height}")
        
        # Save preprocessed image for debugging
        if self.debug_mode:
            processed_path = os.path.join(self.output_dir, f"page_{page_num:03d}_processed.png")
            image.save(processed_path)
            self.logger.info(f"Saved preprocessed image: {processed_path}")
        
        return image
    
    def extract_text_from_image_multi_config(self, image: Image.Image, page_num: int) -> str:
        """Extract text using multiple OCR configurations"""
        
        # Preprocess image
        processed_image = self.preprocess_image_advanced(image, page_num)
        
        # Multiple OCR configurations to try
        ocr_configs = [
            # Standard configuration
            r'--oem 3 --psm 6',
            # Better for tables and structured data
            r'--oem 3 --psm 4',
            # Single uniform block of text
            r'--oem 3 --psm 8',
            # Treat image as single text line
            r'--oem 3 --psm 7',
            # Single word
            r'--oem 3 --psm 13'
        ]
        
        best_text = ""
        best_confidence = 0
        
        for i, config in enumerate(ocr_configs):
            try:
                self.logger.info(f"Trying OCR config {i+1}: {config}")
                
                # Extract text
                text = pytesseract.image_to_string(processed_image, config=config, lang='eng')
                
                # Get confidence data
                try:
                    data = pytesseract.image_to_data(processed_image, config=config, lang='eng', output_type=pytesseract.Output.DICT)
                    confidences = [int(conf) for conf in data['conf'] if int(conf) > 0]
                    avg_confidence = sum(confidences) / len(confidences) if confidences else 0
                except:
                    avg_confidence = len(text)  # Fallback to text length
                
                self.logger.info(f"Config {i+1} - Text length: {len(text)}, Confidence: {avg_confidence:.1f}")
                
                # Keep the text with highest confidence/length
                if avg_confidence > best_confidence:
                    best_confidence = avg_confidence
                    best_text = text
                    
            except Exception as e:
                self.logger.warning(f"OCR config {i+1} failed: {e}")
                continue
        
        # Save raw OCR text for debugging
        if self.debug_mode and best_text:
            text_path = os.path.join(self.output_dir, f"page_{page_num:03d}_raw_ocr.txt")
            with open(text_path, 'w', encoding='utf-8') as f:
                f.write(f"OCR Confidence: {best_confidence:.1f}\n")
                f.write("="*50 + "\n")
                f.write(best_text)
            self.logger.info(f"Saved raw OCR text: {text_path}")
        
        return best_text
    
    def clean_ocr_text_advanced(self, text: str) -> str:
        """Advanced OCR text cleaning and correction"""
        
        if not text:
            return ""
        
        # Common OCR corrections - expanded list
        ocr_corrections = {
            # Name field corrections
            'Nanre': 'Name', 'Narne': 'Name', 'Natne': 'Name', 'Nanie': 'Name',
            'Narme': 'Name', 'Namre': 'Name', 'Nanne': 'Name',
            
            # Father/Husband corrections
            'Fathers': 'Father', 'Fathars': 'Father', 'Fathar': 'Father',
            'Husbands': 'Husband', 'Hurband': 'Husband', 'Husbamd': 'Husband',
            'Hursband': 'Husband', 'Husbanc': 'Husband',
            
            # Age corrections
            'Aqe': 'Age', 'Agg': 'Age', 'Agge': 'Age', 'Agae': 'Age',
            
            # Gender corrections
            'Gendsr': 'Gender', 'Gendet': 'Gender', 'Gencer': 'Gender',
            'Malg': 'Male', 'Malle': 'Male', 'Maie': 'Male',
            'Fernale': 'Female', 'Femala': 'Female', 'Femsle': 'Female',
            'Femaie': 'Female', 'Fenale': 'Female',
            
            # House corrections
            'Houre': 'House', 'Housr': 'House', 'Hourse': 'House',
            'Numbsr': 'Number', 'Numbef': 'Number', 'Numbar': 'Number',
            
            # Common character corrections
            '|': 'I', 'rn': 'm', '0': 'O',  # Only in names context
            'cl': 'd', 'ii': 'll', 'vv': 'w'
        }
        
        # Apply corrections
        for wrong, correct in ocr_corrections.items():
            text = re.sub(r'\b' + re.escape(wrong) + r'\b', correct, text, flags=re.IGNORECASE)
        
        # Clean up whitespace and formatting
        text = re.sub(r'\s+', ' ', text)  # Multiple spaces to single
        text = re.sub(r'\n\s*\n', '\n', text)  # Multiple newlines to single
        text = text.strip()
        
        # Fix common punctuation issues
        text = re.sub(r'\s*:\s*', ': ', text)  # Fix colon spacing
        text = re.sub(r'\s*,\s*', ', ', text)  # Fix comma spacing
        
        return text
    
    def extract_metadata_flexible(self, text: str) -> Dict[str, str]:
        """Flexible metadata extraction with multiple patterns"""
        
        metadata = {}
        text_clean = self.clean_ocr_text_advanced(text)
        
        # Assembly constituency patterns - more flexible
        constituency_patterns = [
            r'Assembly\s+Constituency.*?(\d+).*?[-–—]\s*([A-Z][A-Z\s\(\)]+?)(?:Part|GENERAL|$)',
            r'Constituency.*?(\d+).*?[-–—]\s*([A-Z][A-Z\s\(\)]+?)(?:Part|GENERAL|$)',
            r'(\d{2,3})\s*[-–—]\s*([A-Z][A-Z\s\(\)]{5,}?)(?:Part|GENERAL|$)',
            r'(\d+)\s*[-–—]\s*([A-Z\s]{8,}?)(?:\s|$)',
        ]
        
        for pattern in constituency_patterns:
            match = re.search(pattern, text_clean, re.IGNORECASE | re.DOTALL)
            if match:
                metadata['vidhan_sabha_constituency_no'] = match.group(1).strip()
                metadata['vidhan_sabha_name'] = match.group(2).strip()
                self.logger.info(f"Found constituency: {match.group(1)} - {match.group(2)}")
                break
        
        # Part number patterns
        part_patterns = [
            r'Part\s+No\.?\s*:?\s*(\d+)',
            r'Part\s+Number\s*:?\s*(\d+)',
            r'Part\s*:?\s*(\d+)',
            r'(\d+)\s*$'  # Number at end of line
        ]
        
        for pattern in part_patterns:
            match = re.search(pattern, text_clean, re.IGNORECASE)
            if match:
                metadata['part_no'] = match.group(1).strip()
                self.logger.info(f"Found part number: {match.group(1)}")
                break
        
        # Parliamentary constituency
        parliament_patterns = [
            r'Parliamentary\s+Constituency.*?(\d+).*?[-–—]\s*([A-Z][A-Z\s\(\)]+)',
            r'Lok\s+Sabha.*?(\d+).*?[-–—]\s*([A-Z][A-Z\s\(\)]+)',
        ]
        
        for pattern in parliament_patterns:
            match = re.search(pattern, text_clean, re.IGNORECASE | re.DOTALL)
            if match:
                metadata['lok_sabha_constituency_no'] = match.group(1).strip()
                metadata['lok_sabha_name'] = match.group(2).strip()
                break
        
        # Location information
        location_patterns = {
            'district': r'District\s*:?\s*([A-Z][A-Z\s]+?)(?:\s*\d|\s*Pin|\s*$)',
            'subdivision': r'Subdivision\s*:?\s*([A-Z][A-Z\s]+?)(?:\s*District|\s*$)',
            'tehsil': r'Tehsil\s*:?\s*([A-Z][A-Z\s]+?)(?:\s*$)',
            'block': r'Block\s*:?\s*([A-Z][A-Z\s]+?)(?:\s*$)',
            'pin_code': r'Pin\s+code\s*:?\s*(\d{6})',
        }
        
        for key, pattern in location_patterns.items():
            match = re.search(pattern, text_clean, re.IGNORECASE)
            if match:
                metadata[key] = match.group(1).strip()
        
        return metadata
    
    def find_voter_records_flexible(self, text: str) -> List[str]:
        """Find potential voter record blocks using multiple strategies"""
        
        text_clean = self.clean_ocr_text_advanced(text)
        voter_blocks = []
        
        # Strategy 1: Look for voter ID patterns as record starters
        id_patterns = [
            r'[A-Z]{2,4}/\d+/\d+/\d+',  # WB/24/161/375141
            r'[A-Z]{3}\d{7,10}',         # SVG2562940
            r'[A-Z]{2,4}\d{7,10}',       # MHD1759844
        ]
        
        # Find all voter ID positions
        id_positions = []
        for pattern in id_patterns:
            for match in re.finditer(pattern, text_clean):
                id_positions.append((match.start(), match.end(), match.group()))
        
        # Sort by position
        id_positions.sort()
        
        if id_positions:
            self.logger.info(f"Found {len(id_positions)} potential voter IDs")
            
            # Extract text blocks between voter IDs
            for i, (start, end, voter_id) in enumerate(id_positions):
                # Define block boundaries
                block_start = max(0, start - 50)  # Include some text before ID
                
                if i < len(id_positions) - 1:
                    # Block ends before next voter ID
                    block_end = id_positions[i + 1][0]
                else:
                    # Last block extends to end or next major section
                    block_end = min(len(text_clean), start + 1000)
                
                block_text = text_clean[block_start:block_end].strip()
                
                if block_text and len(block_text) > 20:  # Minimum block size
                    voter_blocks.append(block_text)
        
        # Strategy 2: Look for "Name:" patterns as backup
        if not voter_blocks:
            self.logger.info("No voter IDs found, trying Name: pattern matching")
            
            # Split by "Name:" and process each chunk
            name_parts = re.split(r'Name\s*:', text_clean, flags=re.IGNORECASE)
            
            for i, part in enumerate(name_parts[1:], 1):  # Skip first empty part
                # Take reasonable chunk after "Name:"
                chunk = f"Name: {part[:500].strip()}"  # Limit chunk size
                
                if len(chunk) > 30 and any(keyword in chunk.lower() for keyword in ['age', 'gender', 'house']):
                    voter_blocks.append(chunk)
        
        # Strategy 3: Serial number based splitting
        if not voter_blocks:
            self.logger.info("Trying serial number pattern matching")
            
            # Look for patterns like "1 Name:" or "123 XYZ123456"
            serial_pattern = r'(\d+)\s+([A-Z]{2,4}[/\d]+|Name\s*:)'
            
            parts = re.split(serial_pattern, text_clean, flags=re.IGNORECASE)
            
            # Recombine parts that start with numbers
            for i in range(1, len(parts), 3):  # Every 3rd element after split
                if i + 2 < len(parts):
                    serial = parts[i]
                    prefix = parts[i + 1]
                    content = parts[i + 2][:800]  # Limit content
                    
                    combined = f"{serial} {prefix} {content}".strip()
                    if len(combined) > 30:
                        voter_blocks.append(combined)
        
        self.logger.info(f"Extracted {len(voter_blocks)} potential voter record blocks")
        return voter_blocks
    
    def parse_voter_record_flexible(self, record_text: str, record_index: int, metadata: Dict[str, str]) -> Optional[Dict[str, Any]]:
        """Parse individual voter record with flexible pattern matching"""
        
        # Clean the record text
        clean_text = self.clean_ocr_text_advanced(record_text)
        
        # Initialize voter data
        voter_data = {
            'lok_sabha_constituency_no': metadata.get('lok_sabha_constituency_no', ''),
            'vidhan_sabha_constituency_no': metadata.get('vidhan_sabha_constituency_no', ''),
            'part_no': metadata.get('part_no', ''),
            'vidhan_sabha_name': metadata.get('vidhan_sabha_name', ''),
            'grampanchayet_name': '',
            'ward_no_name': '',
            'tehsil_block_name': metadata.get('tehsil', metadata.get('subdivision', '')),
            'region': self.determine_region(metadata.get('district', '')),
            'division': metadata.get('subdivision', ''),
            'district': metadata.get('district', ''),
            'epic': '',
            'serial_no': '',
            'name': '',
            'father_husband_name': '',
            'house_no': '',
            'age': 0,
            'gender': '',
            'religion': '',
            'caste': '',
            'age_group': '',
            'family_size': ''
        }
        
        # Extract voter ID
        id_patterns = [
            r'([A-Z]{2,4}/\d+/\d+/\d+)',
            r'([A-Z]{3}\d{7,10})',
            r'([A-Z]{2,4}\d{7,10})'
        ]
        
        for pattern in id_patterns:
            match = re.search(pattern, clean_text)
            if match:
                voter_data['epic'] = match.group(1)
                break
        
        # Extract serial number
        serial_match = re.search(r'^(\d+)\s+', clean_text)
        if serial_match:
            voter_data['serial_no'] = serial_match.group(1)
        else:
            voter_data['serial_no'] = str(record_index)
        
        # Extract name - multiple patterns
        name_patterns = [
            r'Name\s*:\s*([A-Za-z\s\.]+?)(?:\s*(?:Father|Husband|Other|Age|House|Gender))',
            r'Name\s*:\s*([A-Za-z\s\.]+?)(?:\s*$)',
            r'(?:^|\s)([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s*(?:Father|Husband|w/o|s/o)',
        ]
        
        for pattern in name_patterns:
            match = re.search(pattern, clean_text, re.IGNORECASE | re.MULTILINE)
            if match:
                name = match.group(1).strip()
                # Validate name (reasonable length, no numbers)
                if 2 <= len(name) <= 50 and not re.search(r'\d', name):
                    voter_data['name'] = name
                    break
        
        # Extract father/husband name
        relation_patterns = [
            r'(?:Father|Husband|Other)(?:s?)?\s*Name\s*:\s*([A-Za-z\s\.]+?)(?:\s*(?:House|Age|Gender|$))',
            r'(?:w/o|s/o|d/o)\s*([A-Za-z\s\.]+?)(?:\s*(?:House|Age|Gender|$))',
            r'(?:Father|Husband)\s*:\s*([A-Za-z\s\.]+?)(?:\s*(?:House|Age|Gender|$))',
        ]
        
        for pattern in relation_patterns:
            match = re.search(pattern, clean_text, re.IGNORECASE | re.MULTILINE)
            if match:
                relation_name = match.group(1).strip()
                if 2 <= len(relation_name) <= 50 and not re.search(r'\d', relation_name):
                    voter_data['father_husband_name'] = relation_name
                    break
        
        # Extract age
        age_patterns = [
            r'Age\s*:\s*(\d+)',
            r'Age\s+(\d+)',
            r'(\d+)\s*years?',
            r'(\d{2})\s*(?:Gender|Male|Female)',
        ]
        
        for pattern in age_patterns:
            match = re.search(pattern, clean_text, re.IGNORECASE)
            if match:
                try:
                    age = int(match.group(1))
                    if 18 <= age <= 120:  # Reasonable age range for voters
                        voter_data['age'] = age
                        voter_data['age_group'] = self.categorize_age(age)
                        break
                except ValueError:
                    continue
        
        # Extract gender
        gender_patterns = [
            r'Gender\s*:\s*(Male|Female|Third\s*Gender)',
            r'(Male|Female)\s*(?:$|\s)',
            r'(?:^|\s)(M|F)\s*(?:$|\s)',
        ]
        
        for pattern in gender_patterns:
            match = re.search(pattern, clean_text, re.IGNORECASE)
            if match:
                gender_text = match.group(1).lower()
                if 'male' in gender_text or gender_text == 'm':
                    voter_data['gender'] = 'M'
                elif 'female' in gender_text or gender_text == 'f':
                    voter_data['gender'] = 'F'
                else:
                    voter_data['gender'] = 'T'
                break
        
        # Extract house number
        house_patterns = [
            r'House\s*(?:Number)?\s*:\s*([A-Za-z0-9/\-\s]+?)(?:\s*(?:Age|Gender|Photo|$))',
            r'(?:^|\s)([A-Za-z0-9/\-]{1,20})\s*(?:Age|Gender)',
            r'(?:Address|House)\s*:\s*([A-Za-z0-9/\-\s]+?)(?:\s*$)',
        ]
        
        for pattern in house_patterns:
            match = re.search(pattern, clean_text, re.IGNORECASE | re.MULTILINE)
            if match:
                house = match.group(1).strip()
                if len(house) <= 30:  # Reasonable house number length
                    voter_data['house_no'] = house
                    break
        
        # Only return record if we have minimum required data
        if voter_data['name'] and voter_data['age'] > 0:
            return voter_data
        
        return None
    
    def categorize_age(self, age: int) -> str:
        """Categorize age into standard groups"""
        if age < 18:
            return 'Under 18'
        elif age <= 29:
            return '18-29'
        elif age <= 45:
            return '30-45'
        else:
            return '45+'
    
    def determine_region(self, district: str) -> str:
        """Determine region/state based on district"""
        if not district:
            return ''
        
        district_lower = district.lower()
        
        # State mappings
        state_mappings = {
            'west bengal': ['hooghly', 'kolkata', 'howrah', 'north 24 parganas', 'south 24 parganas', 'darjeeling'],
            'uttar pradesh': ['agra', 'lucknow', 'kanpur', 'allahabad', 'varanasi', 'meerut', 'ghaziabad'],
            'maharashtra': ['mumbai', 'pune', 'nagpur', 'thane', 'nashik'],
            'gujarat': ['ahmedabad', 'surat', 'vadodara', 'rajkot'],
            'rajasthan': ['jaipur', 'jodhpur', 'udaipur', 'bikaner'],
            'bihar': ['patna', 'gaya', 'muzaffarpur', 'bhagalpur'],
            'odisha': ['bhubaneswar', 'cuttack', 'berhampur', 'sambalpur'],
        }
        
        for state, districts in state_mappings.items():
            if any(d in district_lower for d in districts):
                return state.title()
        
        return district
    
    def extract_from_scanned_pdf(self, pdf_path: str, dpi: int = 300) -> List[Dict[str, Any]]:
        """Main extraction method with comprehensive processing"""
        
        if not OCR_AVAILABLE:
            raise ImportError("OCR libraries not available")
        
        self.logger.info(f"🔍 Starting enhanced OCR extraction from: {pdf_path}")
        
        # Convert PDF to images
        images = self.pdf_to_images(pdf_path, dpi=dpi)
        
        all_voters = []
        combined_text = ""
        
        # Process each page
        for page_num, image in enumerate(images, 1):
            self.logger.info(f"📄 Processing page {page_num}/{len(images)} with enhanced OCR...")
            
            # Extract text using multiple OCR configurations
            page_text = self.extract_text_from_image_multi_config(image, page_num)
            
            if not page_text.strip():
                self.logger.warning(f"⚠️ No text extracted from page {page_num}")
                continue
            
            combined_text += page_text + "\n\n"
            
            # Extract metadata from first few pages
            if page_num <= 3:
                page_metadata = self.extract_metadata_flexible(page_text)
                self.metadata.update(page_metadata)
            
            # Find and parse voter records from this page
            voter_blocks = self.find_voter_records_flexible(page_text)
            
            self.logger.info(f"📊 Page {page_num}: Found {len(voter_blocks)} potential voter blocks")
            
            # Parse each voter block
            for block_index, block in enumerate(voter_blocks):
                voter = self.parse_voter_record_flexible(
                    block, 
                    len(all_voters) + block_index + 1, 
                    self.metadata
                )
                
                if voter:
                    all_voters.append(voter)
            
            self.logger.info(f"✅ Page {page_num}: Extracted {len([v for v in voter_blocks if v])} valid voter records")
        
        # Save combined text for debugging
        if self.debug_mode:
            combined_path = os.path.join(self.output_dir, "all_pages_combined_text.txt")
            with open(combined_path, 'w', encoding='utf-8') as f:
                f.write(f"Metadata extracted:\n{json.dumps(self.metadata, indent=2)}\n")
                f.write("="*80 + "\n")
                f.write(combined_text)
            self.logger.info(f"Saved combined text: {combined_path}")
        
        # Save extraction summary
        if self.debug_mode:
            summary_path = os.path.join(self.output_dir, "extraction_summary.json")
            summary = {
                'timestamp': datetime.now().isoformat(),
                'total_pages': len(images),
                'total_voters_found': len(all_voters),
                'metadata_extracted': self.metadata,
                'voters_per_page': {},
                'sample_voters': all_voters[:3] if all_voters else []
            }
            
            with open(summary_path, 'w', encoding='utf-8') as f:
                json.dump(summary, f, indent=2, ensure_ascii=False)
            self.logger.info(f"Saved extraction summary: {summary_path}")
        
        self.logger.info(f"🎉 Total voters extracted: {len(all_voters)}")
        self.logger.info(f"📋 Metadata found: {self.metadata}")
        
        return all_voters
    
    def save_to_excel(self, voters: List[Dict[str, Any]], output_path: str):
        """Save voter data to Excel with comprehensive formatting"""
        
        if not voters:
            raise Exception("No voter data to save")
        
        # Create DataFrame
        df = pd.DataFrame(voters)
        
        # Standard column mapping
        column_mapping = {
            'lok_sabha_constituency_no': 'Lok Sabha Constinuency No',
            'vidhan_sabha_constituency_no': 'Vidhan Sabha Constituency No',
            'part_no': 'Part No',
            'vidhan_sabha_name': 'Vidhan Sabha Name',
            'grampanchayet_name': 'Grampanchayet Name',
            'ward_no_name': 'Ward No/Ward Name',
            'tehsil_block_name': 'Tehsil/Block Name',
            'region': 'Region',
            'division': 'Division',
            'district': 'District',
            'epic': 'EPIC',
            'serial_no': 'Serial No',
            'name': 'Name',
            'father_husband_name': 'Father/Husband Name',
            'house_no': 'House No',
            'age': 'Age',
            'gender': 'Gender',
            'religion': 'Religion',
            'caste': 'Caste',
            'age_group': 'Age Group',
            'family_size': 'Family Size'
        }
        
        df = df.rename(columns=column_mapping)
        
        # Save to Excel with multiple sheets
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            
            # Dashboard sheet with metadata and statistics
            dashboard_data = []
            dashboard_data.append(['Field', 'Value'])
            dashboard_data.append(['Extraction Method', 'Enhanced OCR Processing'])
            dashboard_data.append(['Processing Date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
            dashboard_data.append(['Total Records', len(voters)])
            
            if self.metadata:
                dashboard_data.append(['', ''])
                dashboard_data.append(['=== CONSTITUENCY INFORMATION ===', ''])
                for key, value in self.metadata.items():
                    dashboard_data.append([key.replace('_', ' ').title(), value])
            
            # Statistics
            if voters:
                male_count = sum(1 for v in voters if v.get('gender') == 'M')
                female_count = sum(1 for v in voters if v.get('gender') == 'F')
                
                dashboard_data.append(['', ''])
                dashboard_data.append(['=== STATISTICS ===', ''])
                dashboard_data.append(['Male Voters', male_count])
                dashboard_data.append(['Female Voters', female_count])
                
                # Age distribution
                age_groups = {}
                for voter in voters:
                    age_group = voter.get('age_group', 'Unknown')
                    age_groups[age_group] = age_groups.get(age_group, 0) + 1
                
                dashboard_data.append(['', ''])
                dashboard_data.append(['=== AGE DISTRIBUTION ===', ''])
                for group, count in age_groups.items():
                    dashboard_data.append([group, count])
            
            dashboard_df = pd.DataFrame(dashboard_data)
            dashboard_df.to_excel(writer, sheet_name='Dashboard', index=False, header=False)
            
            # Main voter data
            df.to_excel(writer, sheet_name='Background Data', index=False)
        
        self.logger.info(f"📁 Excel file saved: {output_path}")
        return True


def main():
    """Main function with enhanced error handling"""
    
    print("🔍 ENHANCED ELECTORAL ROLL OCR CONVERTER")
    print("=" * 60)
    print("📄 Processes scanned PDFs with advanced OCR and debugging")
    print()
    
    if len(sys.argv) != 3:
        print("Usage: python enhanced_electoral_roll_ocr.py input_file.pdf output_file.xlsx")
        print("\nExample: python enhanced_electoral_roll_ocr.py scanned_roll.pdf converted_data.xlsx")
        print("\n🔧 This enhanced version includes:")
        print("• Raw OCR text debugging output")
        print("• Multiple OCR processing strategies") 
        print("• Advanced text cleaning and error correction")
        print("• Flexible voter record detection")
        print("• Comprehensive logging")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    # Validation
    if not os.path.exists(input_file):
        print(f"❌ Error: Input file '{input_file}' not found.")
        sys.exit(1)
    
    if not input_file.lower().endswith('.pdf'):
        print("❌ Error: Input file must be a PDF file.")
        sys.exit(1)
    
    if not output_file.lower().endswith('.xlsx'):
        print("❌ Error: Output file must be an Excel file (.xlsx).")
        sys.exit(1)
    
    if not OCR_AVAILABLE:
        print("❌ OCR libraries not available!")
        print("Install: pip install pytesseract Pillow pdf2image")
        print("\n📥 Also install Tesseract OCR:")
        print("Windows: https://github.com/UB-Mannheim/tesseract/wiki")
        print("macOS: brew install tesseract")
        print("Linux: sudo apt-get install tesseract-ocr")
        sys.exit(1)
    
    print(f"📄 Input: {input_file}")
    print(f"📊 Output: {output_file}")
    print(f"🔧 Debug files will be saved to: ocr_debug_output/")
    print()
    
    # Create enhanced extractor
    extractor = EnhancedElectoralRollExtractor(debug_mode=True)
    
    try:
        print("🚀 Starting enhanced OCR processing...")
        print("⏳ This may take several minutes depending on PDF size...")
        print("🔍 Check ocr_debug_output/ folder for detailed processing info")
        print()
        
        # Process with default DPI (300)
        voters = extractor.extract_from_scanned_pdf(input_file, dpi=300)
        
        if voters:
            # Save to Excel
            extractor.save_to_excel(voters, output_file)
            
            print("\n" + "="*60)
            print("🎉 SUCCESS: Enhanced OCR extraction completed!")
            print(f"📊 Total voters extracted: {len(voters)}")
            print(f"📁 Excel file saved: {output_file}")
            print(f"🔍 Debug files saved in: ocr_debug_output/")
            
            # Display statistics
            if voters:
                male_count = sum(1 for v in voters if v.get('gender') == 'M')
                female_count = sum(1 for v in voters if v.get('gender') == 'F')
                
                print(f"\n📈 Quick Statistics:")
                print(f"   👨 Male voters: {male_count}")
                print(f"   👩 Female voters: {female_count}")
                
                age_groups = {}
                for voter in voters:
                    age_group = voter.get('age_group', 'Unknown')
                    age_groups[age_group] = age_groups.get(age_group, 0) + 1
                
                print(f"   📊 Age distribution:")
                for group, count in age_groups.items():
                    print(f"      {group}: {count}")
        
        else:
            print("\n" + "="*60)
            print("⚠️ WARNING: No voter records were extracted")
            print("\n🔧 Troubleshooting steps:")
            print("1. Check debug files in ocr_debug_output/ folder")
            print("2. Review page_XXX_raw_ocr.txt files for OCR quality")
            print("3. Try increasing DPI or improving PDF scan quality")
            print("4. Verify PDF contains actual voter roll data")
            print("\n📧 If issues persist, share the raw OCR text files for analysis")
        
    except Exception as e:
        print(f"\n❌ Error during processing: {e}")
        print(f"📝 Check electoral_ocr.log for detailed error information")
        print(f"🔍 Debug files may be available in ocr_debug_output/ folder")
        sys.exit(1)


if __name__ == "__main__":
    main()
'''

# Save the improved OCR program
with open('enhanced_electoral_roll_ocr.py', 'w', encoding='utf-8') as f:
    f.write(improved_ocr_program)

print("✅ Enhanced OCR Electoral Roll Converter created!")
print("- enhanced_electoral_roll_ocr.py (Improved OCR with debugging)")