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

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Set Up Google Drive API Credentials

To access private Google Drive files, you need to create OAuth 2.0 credentials:

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the **Google Drive API**:
   - Go to "APIs & Services" > "Library"
   - Search for "Google Drive API"
   - Click "Enable"
4. Create OAuth 2.0 credentials:
   - Go to "APIs & Services" > "Credentials"
### Browser-based Method (Recommended for beginners)

Run the browser-based script:

```bash
python main_browser.py
```

A browser will open. Log in to your Google account if needed, then the file will download automatically.

### API-based Method

Run the APIick "Create Credentials" > "OAuth client ID"
   - Select "Desktop app" as the application type
   - Give it a name (e.g., "Drive Downloader")
   - Click "Create"
5. Download the credentials file:
   - Click the download button next to your newly created OAuth client
   - Save the file as `credentials.json` in the same directory as `main.py`

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
- Participants: Marko Horvat, Ana Kovačić

Output:
- Image: `142_NYC_TeamBuilding_PaidBy_JohnDoe.jpg`
- PDF: `142_NYC_TeamBuilding_PaidBy_JohnDoe.pdf` (with "Marko Horvat, Ana Kovačić" as header)

## Notes

- Browser stays open during batch downloads (login once, download all)
- Files are only downloaded once (checks for existing files)
- PDFs are recreated if missing, even when images exist
- Special characters in headers are fully supported
