# Excel to Markdown Converter

This tool reads Excel workbooks with multiple sheets and embedded images, converting them to Markdown format with extracted images.

## Usage

```bash
python3 process_excel.py
```

The script will:
1. Read `data/Fruits.xlsx`
2. Extract all embedded images to `media/` folder
3. Generate `ExcelFile.md` with sheet data and image links
4. Generate `Report.md` with processing details

## Requirements

- Python 3.6+
- openpyxl library (`pip install openpyxl`)

## Features

- **Robust Image Extraction**: Uses ZIP-based approach to extract all images from `xl/media/*`
- **Complete Sheet Processing**: Converts all sheets to Markdown tables
- **Error Handling**: Reports any issues in `Report.md`
- **No External Dependencies**: Works without relying on openpyxl's `_images` property

## Output Files

- `ExcelFile.md`: Main output with sheet summaries, data tables, and image links
- `Report.md`: Processing report with method details and limitations
- `media/`: Directory containing all extracted images

## Implementation Details

The tool treats the Excel file (.xlsx) as a ZIP archive and directly extracts images from the internal structure, ensuring all embedded images are captured even if they're not accessible through standard openpyxl methods.