# Google Drive Image Downloader & Barcode Generator

A Python script with two main features:
1. Download images from Google Drive with custom naming and PDF generation
2. Generate Code128 barcode images from Excel data

## Features

### Feature 1: Google Drive Image Downloader
- Opens a browser where you manually log in to Google
- Batch downloads images from Google Drive URLs
- Creates PDFs with custom headers
- No OAuth setup required

### Feature 2: Barcode Generator
- Generates Code128 barcodes from Excel file
- Saves barcodes as PNG images with text below
- Customizable text size
- Automatic file naming from Excel data

## Setup

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Browser-based Method (Recommended for beginners)

Run the browser-based script:

```bash
python main_browser.py
```

A browser will open. Log in to your Google account if needed, then the file will download automatically.

## Usage

### Google Drive Image Downloader

Run the script:

```bash
python3 main_browser.py
```

Or with a filename argument:

```bash
python3 main_browser.py july.xlsx
```

#### Excel Format

Expected columns (no header):
- Column 1: id
- Column 2: url (Google Drive URL)
- Column 3: participants
- Column 4: name
- Column 5: event
- Column 6: place

### What Happens

1. **First run with new files**:
   - Browser opens (log in to Google once)
   - Downloads all images to `downloaded_images/`
   - Creates PDFs in `pdf_images/`
   - Browser closes

2. **Subsequent runs**:
   - Skips already downloaded images
   - Only opens browser if new files need downloading
   - Creates any missing PDFs

### Output

**Images**: Saved to `downloaded_images/` as:
```
{id}_{place}_{event}_PaidBy_{name}.jpg
```

**PDFs**: Saved to `pdf_images/` with same naming but:
- `.pdf` extension
- Header text from participants column (large font, 48pt)
- Supports special characters (č, š, ž, đ)


### Barcode Generator

### Google Drive Downloads
- Browser stays open during batch downloads (login once, download all)
- Files are only downloaded once (checks for existing files)
- PDFs are recreated if missing, even when images exist
- Special characters in headers are fully supported

### Barcode Generation
- Automatically detects `Barcodes.xlsx` and switches to barcode mode
- Code128 format supports letters, numbers, and special characters
- Existing barcodes are skipped (won't regenerate)
- Invalid or empty codes are skipped with error message
```

#### Excel Format

Expected columns (no header):
- Column 1: name (used as filename)
- Column 2: code (barcode value)

#### What Happens

1. Reads Excel file with name and code columns
2. Generates Code128 barcodes with text below
3. Saves as PNG images to `barcodes/` folder
4. Skips existing files automatically

#### Output

**Barcodes**: Saved to `barcodes/` as:
```
{name}.png
```

Example:
- If name = "Product_ABC" and code = "123456789"
- Output: `barcodes/Product_ABC.png` (barcode image with "123456789" text below)

#### Barcode Settings

- Font size: 7pt (compact text)
- Text distance: 4px below barcode
- Format: Code128 (supports alphanumeric)
- Image format: PNG
### Example

For a row with:
- ID: 142
- Place: NYC
- Event: TeamBuilding
- Name: John Doe
- Participants: John Snow, Lionel Messi

Output:
- Image: `142_NYC_TeamBuilding_PaidBy_JohnDoe.jpg`
- PDF: `142_NYC_TeamBuilding_PaidBy_JohnDoe.pdf` (with "John Snow, Lionel Messi" as header)

## Notes

- Browser stays open during batch downloads (login once, download all)
- Files are only downloaded once (checks for existing files)
- PDFs are recreated if missing, even when images exist
- Special characters in headers are fully supported
