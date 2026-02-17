#!/usr/bin/env python3
"""
Google Drive Image Downloader (Browser-based)
Downloads images from Google Drive by opening a browser for authentication
Reads URLs from Excel file and downloads with custom naming
"""

import os
import time
import re
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import pandas as pd
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import barcode
from barcode.writer import ImageWriter


def extract_file_id(url):
    """
    Extract file ID from various Google Drive URL formats.
    """
    patterns = [
        r'/d/([a-zA-Z0-9-_]+)',
        r'id=([a-zA-Z0-9-_]+)',
        r'^([a-zA-Z0-9-_]+)$'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    
    return None


def create_pdf_from_image(image_path, pdf_path, header_text):
    """
    Create a PDF from an image with header text above it.
    Supports UTF-8 characters including č, š, ž, đ.
    
    Args:
        image_path: Path to the source image
        pdf_path: Path where to save the PDF
        header_text: Text to display above the image
    """
    try:
        # Open the image to get dimensions
        img = Image.open(image_path)
        img_width, img_height = img.size
        
        # Calculate PDF dimensions (adding space for header)
        header_height = 1.5 * inch  # Space for header text (increased for larger font)
        pdf_width = img_width
        pdf_height = img_height + header_height
        
        # Create PDF with custom page size
        c = canvas.Canvas(pdf_path, pagesize=(pdf_width, pdf_height))
        
        # Try to use Arial Unicode font which supports extended characters
        try:
            pdfmetrics.registerFont(TTFont('ArialUnicode', '/System/Library/Fonts/Supplemental/Arial Unicode.ttf'))
            c.setFont("ArialUnicode", 48)
        except:
            # Fallback to Helvetica
            c.setFont("Helvetica-Bold", 48)
        
        text_y = pdf_height - 80  # Position near top (adjusted for larger font)
        c.drawCentredString(pdf_width / 2, text_y, header_text)
        
        # Add image below the text
        c.drawImage(image_path, 0, 0, width=img_width, height=img_height)
        
        # Save PDF
        c.save()
        return True
        
    except Exception as e:
        print(f"  ✗ Error creating PDF: {str(e)}")
        return False


def sanitize_filename(text):
    """Remove or replace characters that are invalid in filenames."""
    if pd.isna(text):
        return ""
    # Convert to string and remove/replace invalid characters
    text = str(text).strip()
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        text = text.replace(char, '-')
    return text


def download_with_browser(driver, drive_url, output_path, download_dir):
    """
    Download a file from Google Drive using browser automation.
    
    Args:
        driver: Existing webdriver instance (reused across downloads)
        drive_url: Google Drive URL or file ID
        output_path: Full path where to save the downloaded file (including .jpg extension)
        download_dir: Directory where Chrome will initially download the file
    """
    file_id = extract_file_id(drive_url)
    if not file_id:
        print("Error: Could not extract file ID from URL")
        return None
    
    # Convert to direct download URL
    file_url = f"https://drive.google.com/file/d/{file_id}/view"
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    
    print(f"File ID: {file_id}")
    print(f"URL: {file_url}")
    
    try:
        # Navigate to Google Drive file
        print(f"Opening: {file_url}")
        driver.get(file_url)
        
        # Wait for the page to load
        wait = WebDriverWait(driver, 60)  # 1 minute timeout
        
        try:
            # Try to find the download button
            print("Looking for download button...")
            download_button = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '[aria-label*="Download"], [aria-label*="download"], button[aria-label*="Download"]'))
            )
            print("Download button found! Clicking...")
            download_button.click()
            
        except:
            # Alternative: try direct download URL
            print("Trying direct download URL...")
            driver.get(download_url)
            
            # Check if there's a confirmation button for large files
            try:
                confirm_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "uc-download-link"))
                )
                confirm_button.click()
            except:
                pass
        
        # Wait for download to complete
        print("Downloading file...")
        time.sleep(3)  # Give it time to start downloading
        
        # Check for downloaded files
        max_wait = 60  # seconds
        elapsed = 0
        downloaded_file = None
        
        while elapsed < max_wait:
            files = os.listdir(download_dir)
            # Look for files that don't end with .crdownload (Chrome's temp download extension)
            completed_files = [f for f in files if not f.endswith('.crdownload') and not f.startswith('.')]
            
            if completed_files:
                # Get the most recently modified file
                latest_file = max(
                    [os.path.join(download_dir, f) for f in completed_files],
                    key=os.path.getmtime
                )
                
                # Check if it was modified recently (within last 60 seconds)
                if time.time() - os.path.getmtime(latest_file) < 60:
                    downloaded_file = latest_file
                    break
            
            time.sleep(1)
            elapsed += 1
        
        if downloaded_file:
            print("✓ File downloaded")
            
            # Move and rename the file (output_path already has .jpg extension)
            if os.path.abspath(downloaded_file) != os.path.abspath(output_path):
                os.rename(downloaded_file, output_path)
            
            print(f"✓ Saved as: {os.path.basename(output_path)}\n")
            return output_path
        else:
            print("✗ Could not detect downloaded file\n")
            return None
            
    except Exception as e:
        print(f"✗ Error: {str(e)}\n")
        driver.quit()
        return None

def process_excel_file_for_barcodes(excel_file, headless=False):
    """
    Process an Excel file specifically for barcodes.
    Expected columns: name, code
    """
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"Error: File '{excel_file}' not found")
        return

    # Read Excel file
    print(f"Reading Excel file: {excel_file}")
    try:
        # Read without header, assign column names manually
        df = pd.read_excel(excel_file, header=None, names=['name', 'code'])
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return None

    # Verify required columns
    required_columns = ['name', 'code']
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        print(f"Error: Missing required columns: {', '.join(missing_columns)}")
        print(f"Available columns: {', '.join(map(str, df.columns))}")
        return

    print(f"Loaded {len(df)} records")
    print(f"Columns: {', '.join(df.columns)}\n")
    
    # Create output directory for barcodes
    output_dir = "barcodes"
    os.makedirs(output_dir, exist_ok=True)
    print(f"Output directory: {output_dir}")
    print("=" * 80)
    
    # Process each row
    successful = 0
    failed = 0
    
    for idx, row in df.iterrows():
        name = sanitize_filename(row['name']) if not pd.isna(row['name']) else f"barcode_{idx}"
        code = str(row['code']).strip() if not pd.isna(row['code']) else ""
        
        print(f"\n[{idx + 1}/{len(df)}] Processing:")
        print(f"  Name: {name}")
        print(f"  Code: {code}")
        
        # Check if code exists
        if not code:
            print("  ✗ Skipping: No code provided")
            failed += 1
            continue
        
        # Create filename
        filename = f"{name}.png"
        output_path = os.path.join(output_dir, filename)
        
        # Check if file already exists
        if os.path.exists(output_path):
            print(f"  ⊘ Skipping: File already exists - {filename}")
            successful += 1
            continue
        
        try:
            # Generate barcode (using Code128 by default, you can change to EAN13, etc.)
            # Code128 is versatile and works with alphanumeric codes
            EAN = barcode.get_barcode_class('code128')
            
            # Configure writer options for smaller text
            writer_options = {
                'font_size': 7,  # Smaller font size (default is 10)
                'text_distance': 4,  # Distance between barcode and text
            }
            
            ean = EAN(code, writer=ImageWriter())
            
            # Save without extension as it will be added automatically
            output_path_no_ext = os.path.join(output_dir, name)
            saved_path = ean.save(output_path_no_ext, options=writer_options)
            
            print(f"  ✓ Barcode saved: {os.path.basename(saved_path)}")
            
            successful += 1
            
        except Exception as e:
            print(f"  ✗ Error generating barcode: {str(e)}")
            failed += 1
    
    # Summary
    print("\n" + "=" * 80)
    print(f"\nBarcode Generation Summary:")
    print(f"  ✓ Successful: {successful}")
    print(f"  ✗ Failed: {failed}")
    print(f"  Total: {len(df)}")
    print(f"\nBarcodes saved to: {output_dir}/")


def process_excel_file(excel_file, headless=False):
    """
    Process an Excel file with Google Drive URLs and download all files.
    
    Expected columns: id, url, participants, name, event, place
    """
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"Error: File '{excel_file}' not found")
        return
    
    # Read Excel file
    print(f"Reading Excel file: {excel_file}")
    try:
        # Read without header, assign column names manually
        df = pd.read_excel(excel_file, header=None, names=['id', 'url', 'participants', 'name', 'event', 'place'])
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        return None

    # Verify required columns
    required_columns = ['id', 'url', 'participants', 'name', 'event', 'place']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"Error: Missing required columns: {', '.join(missing_columns)}")
        print(f"Available columns: {', '.join(map(str, df.columns))}")
        return
    
    print(f"Loaded {len(df)} records")
    print(f"Columns: {', '.join(df.columns)}\n")
    
    # Create output directory
    output_dir = "downloaded_images"
    os.makedirs(output_dir, exist_ok=True)
    print(f"Output directory: {output_dir}")
    
    # Create PDF output directory
    pdf_dir = "pdf_images"
    os.makedirs(pdf_dir, exist_ok=True)
    print(f"PDF directory: {pdf_dir}")
    
    # Create temp download directory (Chrome will download here first)
    temp_download_dir = os.path.abspath(output_dir)
    
    # Set up Chrome options
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless")
    
    prefs = {
        "download.default_directory": temp_download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Start browser once for all downloads
    print("\nStarting browser...")
    driver = webdriver.Chrome(options=chrome_options)
    
    print("Please log in to Google in the browser window...")
    print("The browser will remain open for all downloads.\n")
    
    print("=" * 80)
    
    # Process each row
    successful = 0
    failed = 0
    last_id = None  # Track last ID for auto-increment
    
    for idx, row in df.iterrows():
        # Handle ID: use previous + 1 if empty
        if pd.isna(row['id']) or str(row['id']).strip() == '':
            if last_id is not None:
                current_id = str(int(float(last_id)) + 1)
            else:
                current_id = "1"
        else:
            current_id = str(int(float(row['id'])))
        
        last_id = current_id
        
        print(f"\n[{idx + 1}/{len(df)}] Processing:")
        print(f"  ID: {current_id}")
        print(f"  Name: {row['name']}")
        print(f"  Event: {row['event']}")
        
        # Check if URL exists
        if pd.isna(row['url']) or not str(row['url']).strip():
            print("  ✗ Skipping: No URL provided")
            failed += 1
            continue
        
        # Create filename: {id}_{place}_{event}_PaidBy_{name}.jpg
        filename = f"{sanitize_filename(current_id)}_{sanitize_filename(row['place'])}_{sanitize_filename(row['event'])}_PaidBy_{sanitize_filename(row['name'])}.jpg"
        
        # Full output path
        output_path = os.path.join(output_dir, filename)
        
        # Check if file already exists
        if os.path.exists(output_path):
            print(f"  ⊘ Skipping: File already exists - {filename}")
            successful += 1  # Count as successful since we have the file
            
            # Still create PDF if it doesn't exist
            pdf_filename = filename.replace('.jpg', '.pdf')
            pdf_path = os.path.join(pdf_dir, pdf_filename)
            
            if not os.path.exists(pdf_path):
                header_text = str(row['participants']) if not pd.isna(row['participants']) else ""
                print(f"  Creating PDF with header: '{header_text}'")
                if create_pdf_from_image(output_path, pdf_path, header_text):
                    print(f"  ✓ PDF saved: {pdf_filename}")
            
            continue
        
        # Download the file
        result = download_with_browser(driver, row['url'], output_path, temp_download_dir)
        
        if result:
            successful += 1
            
            # Create PDF from the downloaded image
            pdf_filename = filename.replace('.jpg', '.pdf')
            pdf_path = os.path.join(pdf_dir, pdf_filename)
            
            # Get header text from participants column (column 3)
            header_text = str(row['participants']) if not pd.isna(row['participants']) else ""
            
            print(f"  Creating PDF with header: '{header_text}'")
            if create_pdf_from_image(output_path, pdf_path, header_text):
                print(f"  ✓ PDF saved: {pdf_filename}")
            
        else:
            failed += 1
    
    # Close the browser if it was opened
    if driver:
        print("\nClosing browser...")
        driver.quit()
    
    # Summary
    print("=" * 80)
    print(f"\nDownload Summary:")
    print(f"  ✓ Successful: {successful}")
    print(f"  ✗ Failed: {failed}")
    print(f"  Total: {len(df)}")
    print(f"\nImages saved to: {output_dir}/")
    print(f"PDFs saved to: {pdf_dir}/")


def main():
    """Main function."""
    print("=" * 80)
    print("Google Drive Batch Downloader (Browser-based) or Barcode Processor")
    print("=" * 80)
    print("\nThis will read URLs from an Excel file and download all images or process barcodes based on the file provided.")
    
    # Check for command-line argument
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
        print(f"Using file from arguments: {excel_file}")
    else:
        # Get Excel filename from user input
        excel_file = input("Enter Excel filename (e.g., July.xlsx): ").strip()
    
    if not excel_file:
        print("Error: No filename provided")
        return
    
    # Process the Excel file
    if excel_file == 'Barcodes.xlsx':
        process_excel_file_for_barcodes(excel_file, True)
    else:
        process_excel_file(excel_file, False)


if __name__ == "__main__":
    main()
