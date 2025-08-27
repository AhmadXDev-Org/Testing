#!/usr/bin/env python3
"""
Excel to Markdown converter with embedded image extraction
Using robust unzip approach to extract images from xl/media/*
"""
import os
import sys
import zipfile
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path
import traceback

class ExcelToMarkdownConverter:
    def __init__(self, excel_path="data/Fruits.xlsx"):
        self.excel_path = excel_path
        self.media_dir = "media"
        self.images_extracted = []
        self.sheet_image_mapping = {}
        self.errors = []
        self.method_used = []
        
    def ensure_directories(self):
        """Create necessary directories"""
        os.makedirs(self.media_dir, exist_ok=True)
        
    def extract_images_unzip_approach(self):
        """
        Extract images using robust unzip approach
        Treat .xlsx as ZIP and extract from xl/media/*
        """
        self.method_used.append("Robust unzip approach")
        
        try:
            with zipfile.ZipFile(self.excel_path, 'r') as zip_ref:
                # List all files in the ZIP
                file_list = zip_ref.namelist()
                
                # Extract images from xl/media/*
                media_files = [f for f in file_list if f.startswith('xl/media/')]
                
                for media_file in media_files:
                    try:
                        # Extract the image
                        image_data = zip_ref.read(media_file)
                        
                        # Get filename from path
                        filename = os.path.basename(media_file)
                        output_path = os.path.join(self.media_dir, filename)
                        
                        # Write image to media directory
                        with open(output_path, 'wb') as f:
                            f.write(image_data)
                            
                        self.images_extracted.append({
                            'original_path': media_file,
                            'output_path': output_path,
                            'filename': filename
                        })
                        
                        print(f"Extracted image: {media_file} -> {output_path}")
                        
                    except Exception as e:
                        error_msg = f"Error extracting {media_file}: {str(e)}"
                        self.errors.append(error_msg)
                        print(error_msg)
                        
                # Try to map images to sheets via xl/drawings/*
                self.map_images_to_sheets(zip_ref, file_list)
                
        except Exception as e:
            error_msg = f"Error in unzip approach: {str(e)}"
            self.errors.append(error_msg)
            print(error_msg)
            
    def map_images_to_sheets(self, zip_ref, file_list):
        """
        Try to map images to sheets by analyzing xl/drawings/* and relationships
        """
        try:
            # Get drawing files
            drawing_files = [f for f in file_list if f.startswith('xl/drawings/')]
            
            for drawing_file in drawing_files:
                try:
                    drawing_data = zip_ref.read(drawing_file)
                    # Parse XML to find image references
                    root = ET.fromstring(drawing_data)
                    
                    # This is a simplified approach - real implementation would need
                    # to parse the complex Excel drawing relationships
                    # For now, we'll note this limitation
                    pass
                    
                except Exception as e:
                    self.errors.append(f"Error parsing {drawing_file}: {str(e)}")
                    
        except Exception as e:
            self.errors.append(f"Error mapping images to sheets: {str(e)}")
            
    def extract_images_library_approach(self):
        """
        Try to extract images using openpyxl-image-loader
        """
        self.method_used.append("Library approach (openpyxl-image-loader)")
        
        try:
            workbook = openpyxl.load_workbook(self.excel_path)
            
            for sheet_name in workbook.sheetnames:
                try:
                    sheet = workbook[sheet_name]
                    image_loader = SheetImageLoader(sheet)
                    
                    # Get images for this sheet
                    sheet_images = []
                    for image in sheet._images:
                        try:
                            # Get image from loader
                            pil_image = image_loader.get(image.anchor._from.row, image.anchor._from.col)
                            if pil_image:
                                # Save image
                                filename = f"sheet_{sheet_name}_image_{len(sheet_images)}.png"
                                filepath = os.path.join(self.media_dir, filename)
                                pil_image.save(filepath)
                                
                                sheet_images.append({
                                    'filename': filename,
                                    'row': image.anchor._from.row,
                                    'col': image.anchor._from.col,
                                    'cell': f"{chr(65 + image.anchor._from.col)}{image.anchor._from.row + 1}"
                                })
                                
                                print(f"Extracted image from sheet '{sheet_name}' at cell {sheet_images[-1]['cell']}")
                                
                        except Exception as e:
                            self.errors.append(f"Error extracting image from sheet {sheet_name}: {str(e)}")
                            
                    if sheet_images:
                        self.sheet_image_mapping[sheet_name] = sheet_images
                        
                except Exception as e:
                    self.errors.append(f"Error processing sheet {sheet_name} for images: {str(e)}")
                    
        except Exception as e:
            error_msg = f"Error in library approach: {str(e)}"
            self.errors.append(error_msg)
            print(error_msg)
            
    def read_excel_sheets(self):
        """
        Read all sheets from Excel file and convert to data structures
        """
        try:
            workbook = openpyxl.load_workbook(self.excel_path, data_only=True)
            sheets_data = {}
            
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                # Get sheet dimensions
                max_row = sheet.max_row
                max_col = sheet.max_column
                
                # Read sheet data
                sheet_data = []
                for row in sheet.iter_rows(values_only=True):
                    # Convert None values to empty strings
                    row_data = [str(cell) if cell is not None else "" for cell in row]
                    sheet_data.append(row_data)
                    
                sheets_data[sheet_name] = {
                    'data': sheet_data,
                    'rows': max_row,
                    'cols': max_col
                }
                
                print(f"Read sheet '{sheet_name}': {max_row} rows, {max_col} columns")
                
            return sheets_data
            
        except Exception as e:
            error_msg = f"Error reading Excel sheets: {str(e)}"
            self.errors.append(error_msg)
            print(error_msg)
            return {}
            
    def convert_to_markdown_table(self, sheet_data):
        """
        Convert sheet data to Markdown table format
        """
        if not sheet_data:
            return "No data available"
            
        markdown_lines = []
        
        # Handle header row
        if sheet_data:
            header = sheet_data[0]
            markdown_lines.append("| " + " | ".join(header) + " |")
            markdown_lines.append("| " + " | ".join(["---"] * len(header)) + " |")
            
            # Handle data rows
            for row in sheet_data[1:]:
                # Pad row to match header length
                padded_row = row + [""] * (len(header) - len(row))
                markdown_lines.append("| " + " | ".join(padded_row[:len(header)]) + " |")
                
        return "\n".join(markdown_lines)
        
    def generate_excel_md(self, sheets_data):
        """
        Generate ExcelFile.md with summary and per-sheet sections
        """
        markdown_content = []
        
        # Summary section
        markdown_content.append("# Excel File Summary\n")
        markdown_content.append("## Summary\n")
        
        for sheet_name, info in sheets_data.items():
            markdown_content.append(f"- **{sheet_name}**: {info['rows']} rows, {info['cols']} columns")
            
        markdown_content.append("\n")
        
        # Per-sheet sections
        for sheet_name, info in sheets_data.items():
            markdown_content.append(f"## Sheet: {sheet_name}\n")
            markdown_content.append(f"Dimensions: {info['rows']} rows × {info['cols']} columns\n")
            markdown_content.append(self.convert_to_markdown_table(info['data']))
            markdown_content.append("\n")
            
        # Images section
        markdown_content.append("## Images\n")
        
        # Per-sheet image mapping
        if self.sheet_image_mapping:
            markdown_content.append("### Images by Sheet\n")
            for sheet_name, images in self.sheet_image_mapping.items():
                markdown_content.append(f"#### {sheet_name}\n")
                for img in images:
                    markdown_content.append(f"- ![Image at {img['cell']}]({img['filename']}) - Located at cell {img['cell']}")
                markdown_content.append("")
        
        # Global image list
        if self.images_extracted:
            markdown_content.append("### All Extracted Images\n")
            for img in self.images_extracted:
                markdown_content.append(f"- ![{img['filename']}]({img['output_path']})")
                
        elif not self.sheet_image_mapping:
            markdown_content.append("No images found in the Excel file.\n")
            
        return "\n".join(markdown_content)
        
    def generate_report_md(self):
        """
        Generate Report.md with process details and any errors
        """
        report_content = []
        
        report_content.append("# Excel Processing Report\n")
        
        # Methods used
        report_content.append("## Methods Used\n")
        for method in self.method_used:
            report_content.append(f"- {method}")
        report_content.append("")
        
        # File processing
        report_content.append("## File Processing\n")
        report_content.append(f"- Source file: {self.excel_path}")
        report_content.append(f"- File exists: {os.path.exists(self.excel_path)}")
        if os.path.exists(self.excel_path):
            report_content.append(f"- File size: {os.path.getsize(self.excel_path)} bytes")
        report_content.append("")
        
        # Images extracted
        report_content.append("## Images Extracted\n")
        report_content.append(f"- Total images extracted: {len(self.images_extracted)}")
        for img in self.images_extracted:
            report_content.append(f"  - {img['original_path']} → {img['output_path']}")
        report_content.append("")
        
        # Limitations
        report_content.append("## Limitations\n")
        report_content.append("- Drawing to sheet mapping is complex and may not be 100% accurate")
        report_content.append("- Some embedded objects might not be detected if they're not in xl/media/")
        if not self.sheet_image_mapping:
            report_content.append("- Unable to map images to specific sheets - images are listed globally")
        report_content.append("")
        
        # Errors
        if self.errors:
            report_content.append("## Errors and Issues\n")
            for error in self.errors:
                report_content.append(f"- {error}")
        else:
            report_content.append("## Status\n")
            report_content.append("- No errors encountered during processing")
            
        return "\n".join(report_content)
        
    def process_excel_file(self):
        """
        Main processing function
        """
        print(f"Processing Excel file: {self.excel_path}")
        
        # Ensure directories exist
        self.ensure_directories()
        
        # Check if file exists
        if not os.path.exists(self.excel_path):
            error_msg = f"Excel file not found: {self.excel_path}"
            self.errors.append(error_msg)
            print(error_msg)
            return False
            
        # Extract images using both approaches
        print("Extracting images using unzip approach...")
        self.extract_images_unzip_approach()
        
        print("Extracting images using library approach...")
        self.extract_images_library_approach()
        
        # Read Excel sheets
        print("Reading Excel sheets...")
        sheets_data = self.read_excel_sheets()
        
        if not sheets_data:
            self.errors.append("No sheet data could be read from Excel file")
            return False
            
        # Generate output files
        print("Generating ExcelFile.md...")
        excel_md_content = self.generate_excel_md(sheets_data)
        with open("ExcelFile.md", "w", encoding="utf-8") as f:
            f.write(excel_md_content)
            
        print("Generating Report.md...")
        report_md_content = self.generate_report_md()
        with open("Report.md", "w", encoding="utf-8") as f:
            f.write(report_md_content)
            
        print(f"Processing complete!")
        print(f"- Extracted {len(self.images_extracted)} images")
        print(f"- Processed {len(sheets_data)} sheets")
        print(f"- Generated ExcelFile.md and Report.md")
        
        return True

def main():
    converter = ExcelToMarkdownConverter()
    success = converter.process_excel_file()
    
    if not success:
        print("Processing failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()