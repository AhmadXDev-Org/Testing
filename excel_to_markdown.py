#!/usr/bin/env python3
"""
Excel to Markdown converter with image extraction
Implements both library and robust unzip approaches for image extraction
"""
import openpyxl
import zipfile
import os
import shutil
from pathlib import Path
import pandas as pd
from openpyxl_image_loader import SheetImageLoader

class ExcelToMarkdownConverter:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.workbook = None
        self.media_dir = Path("media")
        self.images_found = []
        self.errors = []
        self.methods_used = []
        
    def load_workbook(self):
        """Load the Excel workbook"""
        try:
            self.workbook = openpyxl.load_workbook(self.excel_path, data_only=True)
            return True
        except Exception as e:
            self.errors.append(f"Failed to load workbook: {e}")
            return False
    
    def extract_images_library_approach(self):
        """Extract images using openpyxl-image-loader"""
        self.methods_used.append("Library approach (openpyxl-image-loader)")
        found_images = {}
        
        if not self.workbook:
            return found_images
            
        for sheet_name in self.workbook.sheetnames:
            ws = self.workbook[sheet_name]
            sheet_images = []
            
            try:
                image_loader = SheetImageLoader(ws)
                
                # Check all cells for images
                for row in range(1, ws.max_row + 1):
                    for col in range(1, ws.max_column + 1):
                        cell_ref = f"{chr(64+col)}{row}"
                        try:
                            image = image_loader.get(cell_ref)
                            if image:
                                # Save the image
                                self.media_dir.mkdir(exist_ok=True)
                                image_filename = f"{sheet_name}_{cell_ref}_image.png"
                                image_path = self.media_dir / image_filename
                                image.save(image_path)
                                sheet_images.append({
                                    'filename': image_filename,
                                    'cell': cell_ref,
                                    'method': 'library'
                                })
                        except Exception:
                            pass  # No image at this cell
                            
                if sheet_images:
                    found_images[sheet_name] = sheet_images
                    
            except Exception as e:
                self.errors.append(f"Error extracting images from {sheet_name} using library: {e}")
        
        return found_images
    
    def extract_images_unzip_approach(self):
        """Extract images using robust unzip approach"""
        self.methods_used.append("Robust unzip approach (treat .xlsx as ZIP)")
        extracted_images = []
        
        try:
            # Create media directory
            self.media_dir.mkdir(exist_ok=True)
            
            with zipfile.ZipFile(self.excel_path, 'r') as zip_ref:
                # Get all files in the ZIP
                all_files = zip_ref.namelist()
                
                # Find media files
                media_files = [f for f in all_files if f.startswith('xl/media/')]
                
                for media_file in media_files:
                    try:
                        # Extract the file
                        file_data = zip_ref.read(media_file)
                        
                        # Get filename from path
                        filename = Path(media_file).name
                        output_path = self.media_dir / filename
                        
                        # Write the file
                        with open(output_path, 'wb') as f:
                            f.write(file_data)
                        
                        extracted_images.append({
                            'filename': filename,
                            'original_path': media_file,
                            'method': 'unzip'
                        })
                        
                    except Exception as e:
                        self.errors.append(f"Error extracting {media_file}: {e}")
                
                # Try to analyze drawing relationships (limited implementation)
                # This is a complex mapping that would require XML parsing
                # For now, we'll note this limitation
                if not media_files:
                    self.errors.append("No media files found in xl/media/ directory")
                else:
                    self.errors.append(f"Found {len(media_files)} media files. Sheet-to-image mapping requires complex XML parsing of drawing relationships and is not implemented in this version.")
                
        except Exception as e:
            self.errors.append(f"Error in unzip approach: {e}")
        
        return extracted_images
    
    def describe_image(self, image_path):
        """Describe what's in an image (placeholder for actual image analysis)"""
        # For now, just return basic info since we can't actually "see" the image
        # In a real implementation, you'd use image recognition here
        return f"Image file: {image_path.name} (format: {image_path.suffix})"
    
    def convert_sheets_to_markdown(self):
        """Convert all sheets to markdown tables"""
        if not self.workbook:
            return ""
        
        markdown_content = []
        
        # Summary section
        markdown_content.append("# Excel File Summary\n")
        markdown_content.append(f"**File:** {self.excel_path}\n")
        markdown_content.append(f"**Sheets:** {len(self.workbook.sheetnames)}\n")
        
        # List sheets with their dimensions
        for sheet_name in self.workbook.sheetnames:
            ws = self.workbook[sheet_name]
            markdown_content.append(f"- **{sheet_name}:** {ws.max_row} rows × {ws.max_column} columns")
        
        markdown_content.append("\n---\n")
        
        # Convert each sheet
        for sheet_name in self.workbook.sheetnames:
            markdown_content.append(f"## Sheet: {sheet_name}\n")
            
            ws = self.workbook[sheet_name]
            
            # Convert to pandas DataFrame for easier markdown conversion
            data = []
            for row in ws.iter_rows(values_only=True):
                data.append(row)
            
            if data:
                # Create DataFrame
                df = pd.DataFrame(data)
                
                # Convert to markdown table
                markdown_table = df.to_markdown(index=False, tablefmt='github')
                markdown_content.append(markdown_table)
            else:
                markdown_content.append("*No data found in this sheet*")
            
            markdown_content.append("\n")
        
        return "\n".join(markdown_content)
    
    def generate_image_section(self, library_images, unzip_images):
        """Generate the images section for markdown"""
        content = []
        content.append("## Images\n")
        
        # Library approach images (per sheet)
        if library_images:
            content.append("### Images found via library approach (per sheet):\n")
            for sheet_name, images in library_images.items():
                content.append(f"**{sheet_name}:**")
                for img in images:
                    content.append(f"- ![{img['filename']}](media/{img['filename']}) at cell {img['cell']}")
                content.append("")
        
        # Unzip approach images (global)
        if unzip_images:
            content.append("### Images extracted via unzip approach:\n")
            for img in unzip_images:
                image_path = self.media_dir / img['filename']
                description = self.describe_image(image_path)
                content.append(f"- ![{img['filename']}](media/{img['filename']}) - {description}")
            content.append("")
        
        if not library_images and not unzip_images:
            content.append("*No images found or extracted*\n")
        
        return "\n".join(content)
    
    def generate_report(self):
        """Generate the Report.md content"""
        content = []
        content.append("# Excel Processing Report\n")
        content.append(f"**File processed:** {self.excel_path}\n")
        content.append(f"**Processing date:** {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        
        content.append("## Methods Used\n")
        for method in self.methods_used:
            content.append(f"- {method}")
        content.append("")
        
        content.append("## File Movement\n")
        content.append(f"- Downloaded Excel file from GitHub attachment to: `{self.excel_path}`")
        content.append("- File is now versioned in the repository\n")
        
        if self.errors:
            content.append("## Errors and Limitations\n")
            for error in self.errors:
                content.append(f"- {error}")
            content.append("")
        
        # Add success metrics
        if self.media_dir.exists():
            image_files = list(self.media_dir.glob("*"))
            content.append(f"## Results\n")
            content.append(f"- Images extracted: {len(image_files)}")
            content.append(f"- Media directory created: `{self.media_dir}`")
            if image_files:
                content.append("- Extracted files:")
                for img_file in image_files:
                    content.append(f"  - {img_file.name}")
        
        return "\n".join(content)
    
    def convert(self):
        """Main conversion method"""
        # Load workbook
        if not self.load_workbook():
            return False
        
        # Extract images using both approaches
        library_images = self.extract_images_library_approach()
        unzip_images = self.extract_images_unzip_approach()
        
        # Convert sheets to markdown
        markdown_content = self.convert_sheets_to_markdown()
        
        # Add images section
        images_section = self.generate_image_section(library_images, unzip_images)
        full_markdown = markdown_content + "\n" + images_section
        
        # Write ExcelFile.md
        with open("ExcelFile.md", "w", encoding="utf-8") as f:
            f.write(full_markdown)
        
        # Generate and write Report.md
        report_content = self.generate_report()
        with open("Report.md", "w", encoding="utf-8") as f:
            f.write(report_content)
        
        return True

def main():
    converter = ExcelToMarkdownConverter("data/Fruits.xlsx")
    success = converter.convert()
    
    if success:
        print("✓ Conversion completed successfully!")
        print("✓ Generated: ExcelFile.md")
        print("✓ Generated: Report.md")
        if converter.media_dir.exists():
            images = list(converter.media_dir.glob("*"))
            print(f"✓ Extracted {len(images)} images to media/ directory")
    else:
        print("✗ Conversion failed!")
        for error in converter.errors:
            print(f"  - {error}")

if __name__ == "__main__":
    main()