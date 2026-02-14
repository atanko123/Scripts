# Google Drive Image Downloader

A Python script to download images (or any files) from Google Drive using authentication.

## Two Methods Available

### Method 1: Browser-based (Simpler) - `main_browser.py`
- Opens a browser where you manually log in to Google
- No OAuth setup required
- More user-friendly

### Method 2: API-based - `main.py`
- Usea. Browser-based Method (Simpler)

Install Chrome browser if not already installed, then run:

```bash
python main_browser.py
```

No additional setup needed! Just log in when the browser opens.

### 2b. API-based Method - Google Drive API with OAuth
- Requires setup of credentials
- Better for automation

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

Run the script:

```bash
python3 main_browser.py
```

Or with a filename argument:

```bash
python3 main_browser.py july.xlsx
```

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
