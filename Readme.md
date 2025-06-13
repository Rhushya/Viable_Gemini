#  Viable Invoice Processing Automation

This Google Apps Script automation processes emails with invoice/bill attachments, extracts key data using AI-powered OCR, saves files to Google Drive, and logs information to Google Sheets.

##  Features

- **Email Processing**: Automatically processes emails with subject "Viable: Trial Document"
- **AI-Powered Data Extraction**: Uses Google Gemini API for intelligent invoice data extraction
- **Multi-OCR Fallback**: Falls back to Google Vision API and Drive OCR when needed
- **Smart File Management**: Saves attachments to Google Drive with structured naming
- **Comprehensive Logging**: Logs all data to Google Sheets with timestamps
- **Automatic Email Management**: Labels and marks emails as processed
- **Flexible Processing**: Can process files from emails or a docs folder
- **Error Handling**: Robust error handling with detailed logging
- **Automated Triggers**: Time-based automation (every 3 hours)

##  Prerequisites

- Google Account with Gmail, Drive, and Sheets access
- Google Apps Script project
- Google Gemini API key (recommended for best results)
- Basic understanding of Google Workspace APIs

##  Quick Setup Guide

### 1. Create Google Resources

#### Google Drive Folders
1. Go to [Google Drive](https://drive.google.com)
2. Create these folders:
   - **"Viable_Test Documents"** (for storing processed files)
   - **"OCR_Temp"** (for temporary OCR processing)
   - **"docs"** (optional - for testing with sample files)
3. Copy the folder IDs from the URLs

#### Google Sheets
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new spreadsheet
3. Copy the sheet ID from the URL

### 2. Set Up Google Apps Script

1. Go to [script.google.com](https://script.google.com)
2. Click **"New Project"**
3. Paste the code from `final_wokring.gs` into the script editor
4. Update the configuration constants at the top:

```javascript
const DRIVE_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';
const SHEET_ID = 'YOUR_SHEET_ID_HERE'; 
const OCR_TEMP_FOLDER_ID = 'YOUR_OCR_TEMP_FOLDER_ID_HERE';
const GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY_HERE'; 
```

### 3. Enable Required APIs

Go to **Resources → Advanced Google Services** and enable:
- Drive API
- Gmail API
- Sheets API

### 4. Get Gemini API Key (Recommended)

1. Go to [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Create a new API key
3. Copy it to the `GEMINI_API_KEY` constant in your script

### 5. Initialize the Script

Run these functions in order:
1. `setupScript()` - Creates necessary sheets and validates setup
2. `testGeminiAPI()` - Tests API connectivity (if using Gemini)
3. `testProcessing()` - Tests the main processing function

### 6. Set Up Automation

Run `setupTrigger()` to create a time-based trigger that runs every 3 hours.

##  Data Output

### Google Sheets Structure

The script creates two sheets automatically:

#### Data Sheet
| Timestamp | Invoice/Bill Date | Invoice/Bill Number | Amount | Vendor/Company Name | Drive File URL | File Type |
|-----------|-------------------|---------------------|---------|---------------------|----------------|-----------|
| 2025-06-12 10:30:00 | 03.06.25 | INV-12345 | Rs 40,000 | Viable Ideas Pvt Ltd | [Link] | PDF |

#### Logs Sheet
| Timestamp | Message |
|-----------|---------|
| 2025-06-12 10:30:00 | ✓ Processing started |
| 2025-06-12 10:30:05 |  Found 2 emails to process |

### File Naming Convention

Files are saved with the format: `Date_Vendor_InvoiceNumber_Amount.extension`

Example: `03.06.25_ViableIdeas_INV12345_Rs40000.pdf`

##  Script Functions

### Main Functions
- `processInvoices()` - Main processing function
- `processAttachment()` - Handles individual files
- `performGeminiExtraction()` - AI-powered data extraction

### Setup & Testing Functions
- `setupScript()` - Initial configuration
- `setupTrigger()` - Creates automation trigger
- `testProcessing()` - Tests main functionality
- `testGeminiAPI()` - Tests API connectivity
- `testDocsFolder()` - Tests folder processing

### OCR Functions
- `performGeminiExtraction()` - Gemini AI extraction (primary)
- `performDriveOCR()` - Google Drive OCR (fallback)
- `fallbackToTraditionalOCR()` - Traditional OCR methods

### Utility Functions
- `getDataSheet()` / `getLogSheet()` - Sheet management
- `logMessage()` - Centralized logging
- `formatFilename()` - File naming

### Sample Data Processing

If no emails are found, the script automatically processes files from the "docs" folder. Upload sample invoices/bills to this folder for testing.

##  Configuration Options

### Supported File Types
- **Images**: JPEG, PNG, GIF, BMP, TIFF, WebP
- **Documents**: PDF
- **Email files**: EML

### OCR Methods (in order of preference)
1. **Google Gemini API** - AI-powered extraction (recommended)
2. **Google Drive OCR** - Built-in Google functionality (free fallback)

### Email Query Customization

Modify the email search query in `processInvoices()`:

```javascript
const query = 'subject:"Viable: Trial Document" -label:Processed';
BUT WE CAN MODIFY THIS PART EITHER TO EXTRACT A PARTICULAR PART OR EXTRACT RECENT MAIL .
```


#  Advanced Features

## Error Handling
- **Comprehensive error logging** to Logs sheet
- **Graceful fallbacks** when APIs fail
- **Automatic retry mechanisms**

## Performance Optimization
- **Efficient batch processing**
- **Automatic cleanup** of temporary files
- **Optimized API usage**

## Security
- API keys stored as **constants** (consider using **Properties Service** for production)
- **File access** restricted to specified folders
- Comprehensive **audit trail**

# Troubleshooting

## Common Issues
### No emails found
- Check email **search query**
- Verify emails have the correct **subject line**
- Use `testDocsFolder()` to process sample files

### OCR not working
- Verify **Gemini API key** is correct
- Check **API quotas** and billing
- Enable required **Google APIs**

### Permission errors
- **Re-authorize** the script
- Check **folder permissions**
- Verify **sheet access**

# Data Extraction Patterns
The script extracts:

- **Invoice Date**: Converted to DD.MM.YY format
- **Vendor Name**: Company that issued the invoice (max 30 chars)
- **Invoice Number**: Invoice/bill/reference number
- **Total Amount**: Final amount in "Rs X,XXX" format

### Extraction Priority
1. Gemini AI analysis (most accurate)
2. Vision API OCR + pattern matching
3. Drive OCR + pattern matching
4. Fallback to email metadata

# 🔧 Production Deployment

## Security Considerations
- Store **API keys** in **PropertiesService** instead of constants
- Implement additional **access controls**
- Set up **monitoring and alerting**

## Scaling
- Adjust **trigger frequency** based on email volume
- Consider **batch processing** for high volumes
- Monitor **API usage and costs**

#  Monitoring & Analytics
- Check the **Logs sheet** regularly for errors
- Monitor **processing success rates**
- Review **API usage and costs**
- Set up **email notifications** for critical errors

## VEDIO DEMO 
