// Configuration - Update these values with your actual IDs
const DRIVE_FOLDER_ID = '1qfst6lndvVJSd_OEfb6rpCUAk0FAKXHk'; 
const SHEET_ID = '1Nj0Sod_l7GFup0B9rR0nlivdlWPRsEcRWVkdsaOr_VU'; 
const OCR_TEMP_FOLDER_ID = '1zE_cbdeK4PCTUu0LaxjseKwEX8nzn5MW'; 

// Google Gemini API Configuration
const GEMINI_API_KEY = 'AIzaSyCPkgaExMLPbJZ9xEL_LHDZ0RVhaANL0_k';
const GEMINI_API_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent';

// Alternative folder name for test documents
const DOCS_FOLDER_NAME = 'docs';

// Main function that processes all invoices from emails
function processInvoices() {
  const logSheet = getLogSheet();
  
  try {
    logMessage('Starting Invoice Processing', logSheet);
    
    // Get all required services and folders ready
    const label = getOrCreateLabel('Processed');
    const folder = getOrCreateDriveFolder('Viable_Test Documents');
    const tempFolder = DriveApp.getFolderById(OCR_TEMP_FOLDER_ID);
    const dataSheet = getDataSheet();
    
    // Search for target emails from the assessment
    const query = 'from:	Tech xyz <tech@viableideas.co> subject:"Viable: Trial Document"';
    const threads = GmailApp.search(query);
    
    // If no emails found, try processing files from docs folder instead
    if (threads.length === 0) {
      logMessage('No matching emails found. Checking docs folder for test files', logSheet);
      return processDocumentsFromFolder(folder, tempFolder, dataSheet, logSheet);
    }

    logMessage(`Found ${threads.length} email thread(s) to process`, logSheet);

    // Process each email thread we found
    let processedCount = 0;
    threads.forEach((thread, threadIndex) => {
      logMessage(`Processing thread ${threadIndex + 1} of ${threads.length}`, logSheet);
      
      // Process each message in the thread
      thread.getMessages().forEach((message, messageIndex) => {
        try {
          const attachments = message.getAttachments();
          
          // If email has no attachments, still mark it as processed
          if (attachments.length === 0) {
            logMessage(`Email ${messageIndex + 1} has no attachments - marking as processed`, logSheet);
            message.markRead();
            try {
              if (typeof message.addLabel === 'function') {
                message.addLabel(label);
              } else {
                thread.addLabel(label);
              }
            } catch (labelError) {
              logMessage(`Could not add label: ${labelError.toString()}`, logSheet);
            }
            return;
          }

          logMessage(`Processing ${attachments.length} attachment(s) from email ${messageIndex + 1}`, logSheet);

          // Process each attachment in the email
          attachments.forEach((attachment, attachIndex) => {
            try {
              if (processAttachment(attachment, message, folder, tempFolder, dataSheet, logSheet)) {
                processedCount++;
                logMessage(`Successfully processed attachment ${attachIndex + 1}: ${attachment.getName()}`, logSheet);
              } else {
                logMessage(`Failed to process attachment ${attachIndex + 1}: ${attachment.getName()}`, logSheet);
              }
            } catch (attachmentError) {
              logMessage(`Error processing attachment: ${attachmentError.toString()}`, logSheet);
            }
          });

          // Mark email as processed after handling all attachments
          message.markRead();
          try {
            if (typeof message.addLabel === 'function') {
              message.addLabel(label);
            } else {
              thread.addLabel(label);
            }
          } catch (labelError) {
            logMessage(`Could not assign label: ${labelError.toString()}`, logSheet);
          }
          
        } catch (messageError) {
          logMessage(`Error processing message: ${messageError.toString()}`, logSheet);
        }
      });
    });

    logMessage(`Processing completed. Successfully processed ${processedCount} attachments`, logSheet);
    
  } catch (error) {
    logMessage(`Main processing error: ${error.toString()}`, logSheet);
  }
}

// Process individual attachment - the core function that handles each file
function processAttachment(attachment, message, folder, tempFolder, dataSheet, logSheet) {
  try {
    const mimeType = attachment.getContentType();
    
    // Check if we support this file type
    if (!isSupportedFileType(mimeType)) {
      logMessage(`Skipped unsupported file type: ${mimeType} for file: ${attachment.getName()}`, logSheet);
      return false;
    }

    logMessage(`Processing file: ${attachment.getName()} (${mimeType})`, logSheet);

    // Try to extract data using Gemini API first
    let extractedData = {};
    
    if (GEMINI_API_KEY && GEMINI_API_KEY.trim() !== '') {
      logMessage('Using Gemini API for data extraction', logSheet);
      extractedData = performGeminiExtraction(attachment, message, logSheet);
      
      // If Gemini fails, fall back to traditional OCR
      if (!extractedData || Object.keys(extractedData).length === 0) {
        logMessage('Gemini API failed, using traditional OCR method', logSheet);
        extractedData = fallbackToTraditionalOCR(attachment, message, tempFolder, logSheet);
      }
    } else {
      logMessage('Gemini API not configured, using traditional OCR', logSheet);
      extractedData = fallbackToTraditionalOCR(attachment, message, tempFolder, logSheet);
    }

    logMessage(`Data extraction completed: ${JSON.stringify(extractedData)}`, logSheet);

    // Create a proper filename based on extracted data
    const fileExtension = getFileExtension(attachment.getName(), mimeType);
    const fileName = formatFilename(extractedData, fileExtension);
    
    logMessage(`Generated filename: ${fileName}`, logSheet);
    
    // Save file to Google Drive
    const file = folder.createFile(attachment);
    file.setName(fileName);
    const fileUrl = file.getUrl();

    logMessage(`File saved to Drive: ${fileName}`, logSheet);

    // Log all data to the Google Sheet as required
    const row = [
      new Date(), // When this was processed
      extractedData.invoiceDate || 'N/A',
      extractedData.invoiceNumber || 'N/A',
      extractedData.totalAmount || 'N/A',
      extractedData.vendorName || 'N/A',
      fileUrl,
      mimeType.split('/')[1].toUpperCase() // File type like PDF, JPG etc
    ];
    
    dataSheet.appendRow(row);
    logMessage(`Data logged to spreadsheet for ${fileName}`, logSheet);

    return true;
    
  } catch (error) {
    logMessage(`Error processing attachment ${attachment.getName()}: ${error.toString()}`, logSheet);
    return false;
  }
}

// Use Gemini API to extract invoice data directly from the file
function performGeminiExtraction(attachment, message, logSheet) {
  try {
    if (!GEMINI_API_KEY || GEMINI_API_KEY.trim() === '') {
      throw new Error('Gemini API key not configured');
    }

    // Convert the file to base64 so we can send it to Gemini
    const blob = attachment.copyBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    const mimeType = attachment.getContentType();
    
    // Create a detailed prompt asking Gemini to extract invoice data
    const prompt = `Analyze this invoice or bill document and extract the following information in JSON format:

{
  "invoiceDate": "DD.MM.YY format (example: 03.06.25)",
  "vendorName": "Company or business name that issued this invoice (maximum 30 characters)",
  "invoiceNumber": "Invoice number, bill number, or reference number",
  "totalAmount": "Final total amount to be paid in format 'Rs X,XXX'"
}

Important instructions:
- For invoiceDate: Convert any date format to DD.MM.YY (example: June 3, 2025 becomes 03.06.25)
- For vendorName: Find the company name that sent the invoice, not the customer name
- For invoiceNumber: Look for invoice number, bill number, or any reference number
- For totalAmount: Find the final amount (look for "Total", "Net Amount", "Final Amount", "Amount Due"). If multiple amounts exist, use the highest one. Format as "Rs X,XXX"
- If any information is missing, use "N/A"
- Return only valid JSON, no extra text

Document to analyze:`;

    // Prepare the request to send to Gemini API
    const requestBody = {
      contents: [{
        parts: [
          {
            text: prompt
          },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Data
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.1, // Keep responses consistent
        maxOutputTokens: 1024
      }
    };

    // Send request to Gemini API
    const response = UrlFetchApp.fetch(`${GEMINI_API_URL}?key=${GEMINI_API_KEY}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());
    
    // Check if the API call was successful
    if (response.getResponseCode() !== 200) {
      throw new Error(`Gemini API error: ${result.error ? result.error.message : 'Unknown error'}`);
    }
    
    // Extract the response text from Gemini
    if (result.candidates && result.candidates[0] && result.candidates[0].content) {
      const responseText = result.candidates[0].content.parts[0].text;
      logMessage(`Gemini response: ${responseText}`, logSheet);
      
      try {
        // Find and parse the JSON from Gemini's response
        const jsonMatch = responseText.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const extractedData = JSON.parse(jsonMatch[0]);
          
          // Clean up the extracted data and add fallbacks
          const cleanedData = {
            invoiceDate: extractedData.invoiceDate || formatDate(message.getDate()),
            vendorName: (extractedData.vendorName || extractSenderName(message)).substring(0, 30),
            invoiceNumber: extractedData.invoiceNumber || generateInvoiceNumber(),
            totalAmount: extractedData.totalAmount || 'N/A'
          };
          
          logMessage('Gemini extraction successful', logSheet);
          return cleanedData;
        } else {
          throw new Error('No JSON found in Gemini response');
        }
      } catch (parseError) {
        logMessage(`JSON parsing error: ${parseError.toString()}`, logSheet);
        throw parseError;
      }
    }
    
    throw new Error('No content in Gemini API response');
    
  } catch (error) {
    logMessage(`Gemini API error: ${error.toString()}`, logSheet);
    return {};
  }
}

// Fallback method using traditional OCR when Gemini fails
function fallbackToTraditionalOCR(attachment, message, tempFolder, logSheet) {
  try {
    logMessage('Starting traditional OCR method', logSheet);
    
    // Use Google Drive's OCR capabilities
    const extractedText = performDriveOCR(attachment, tempFolder, logSheet);

    if (!extractedText || extractedText.trim().length === 0) {
      logMessage(`No text could be extracted from: ${attachment.getName()}`, logSheet);
      // Return basic data when OCR fails completely
      return {
        invoiceDate: formatDate(message.getDate()),
        vendorName: extractSenderName(message),
        invoiceNumber: generateInvoiceNumber(),
        totalAmount: 'N/A'
      };
    }

    // Extract structured data from the OCR text using pattern matching
    const extractedData = extractInvoiceDataFromText(extractedText, message, logSheet);
    
    return extractedData;
    
  } catch (error) {
    logMessage(`Traditional OCR error: ${error.toString()}`, logSheet);
    return {
      invoiceDate: formatDate(message.getDate()),
      vendorName: extractSenderName(message),
      invoiceNumber: generateInvoiceNumber(),
      totalAmount: 'N/A'
    };
  }
}

// Use Google Drive to extract text from files (OCR)
function performDriveOCR(attachment, tempFolder, logSheet) {
  try {
    logMessage(`Starting Drive OCR for: ${attachment.getName()}`, logSheet);
    
    // Create a temporary file in Drive for OCR processing
    const tempFile = tempFolder.createFile(attachment);
    let ocrText = '';
    
    try {
      // Convert to Google Doc to trigger OCR
      const ocrDoc = Drive.Files.copy({
        title: tempFile.getName() + '_ocr_' + Date.now(),
        mimeType: 'application/vnd.google-apps.document'
      }, tempFile.getId());
      
      // Wait for Google to process the OCR
      Utilities.sleep(2000);
      
      // Get the text from the converted document
      const doc = DocumentApp.openById(ocrDoc.id);
      ocrText = doc.getBody().getText();
      
      // Clean up the temporary OCR document
      DriveApp.getFileById(ocrDoc.id).setTrashed(true);
      
    } catch (driveError) {
      logMessage(`Drive OCR conversion error: ${driveError.toString()}`, logSheet);
    }
    
    // Clean up the temporary file
    DriveApp.getFileById(tempFile.getId()).setTrashed(true);
    
    logMessage(`Drive OCR completed. Text length: ${ocrText.length}`, logSheet);
    return ocrText;
    
  } catch (error) {
    logMessage(`Drive OCR error: ${error.toString()}`, logSheet);
    return '';
  }
}

// Process documents from docs folder when no emails are found
function processDocumentsFromFolder(targetFolder, tempFolder, dataSheet, logSheet) {
  try {
    logMessage('Processing documents from docs folder', logSheet);
    
    // Find the docs folder or create it if it doesn't exist
    const docsFolders = DriveApp.getFoldersByName(DOCS_FOLDER_NAME);
    
    if (!docsFolders.hasNext()) {
      logMessage('No docs folder found. Please create a "docs" folder and upload test documents', logSheet);
      return;
    }
    
    const docsFolder = docsFolders.next();
    const files = docsFolder.getFiles();
    
    let processedCount = 0;
    
    // Process each file in the docs folder
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getBlob().getContentType();
      
      if (isSupportedFileType(mimeType)) {
        logMessage(`Processing file from docs: ${file.getName()}`, logSheet);
        
        // Create a fake email object for processing
        const mockEmail = {
          getDate: () => new Date(),
          getFrom: () => 'tech@viableideas.co <tech@viableideas.co>'
        };
        
        if (processFileAsAttachment(file, mockEmail, targetFolder, tempFolder, dataSheet, logSheet)) {
          processedCount++;
        }
      } else {
        logMessage(`Skipping unsupported file: ${file.getName()} (${mimeType})`, logSheet);
      }
    }
    
    logMessage(`Processed ${processedCount} files from docs folder`, logSheet);
    
  } catch (error) {
    logMessage(`Docs folder processing error: ${error.toString()}`, logSheet);
  }
}

// Convert a Drive file to an attachment-like object for processing
function processFileAsAttachment(file, mockEmail, folder, tempFolder, dataSheet, logSheet) {
  try {
    const mimeType = file.getBlob().getContentType();
    const blob = file.getBlob();
    
    // Create an object that looks like an email attachment
    const mockAttachment = {
      getName: () => file.getName(),
      getContentType: () => mimeType,
      copyBlob: () => blob
    };
    
    return processAttachment(mockAttachment, mockEmail, folder, tempFolder, dataSheet, logSheet);
    
  } catch (error) {
    logMessage(`File processing error: ${error.toString()}`, logSheet);
    return false;
  }
}

// Extract invoice data from OCR text using pattern matching
function extractInvoiceDataFromText(text, email, logSheet) {
  logMessage('Starting data extraction from OCR text', logSheet);
  
  // Clean up the text to make pattern matching easier
  const cleanText = text.replace(/\s+/g, ' ').trim();
  
  // Extract each field using different methods
  const invoiceDate = extractDateFromText(cleanText) || formatDate(email.getDate());
  const vendorName = extractVendorNameFromText(cleanText) || extractSenderName(email);
  const invoiceNumber = extractInvoiceNumberFromText(cleanText) || generateInvoiceNumber();
  const totalAmount = extractTotalAmountFromText(cleanText);
  
  const extractedData = {
    invoiceDate: invoiceDate,
    vendorName: vendorName,
    invoiceNumber: invoiceNumber,
    totalAmount: totalAmount
  };
  
  logMessage(`Extracted from text: ${JSON.stringify(extractedData)}`, logSheet);
  return extractedData;
}

// Find and extract amounts from text, prioritizing final/total amounts
function extractTotalAmountFromText(text) {
  // Different patterns to find amounts in the text
  const amountPatterns = [
    // Look for total/final/net with currency
    /(?:total|final|net|grand\s*total|amount\s*due|payable|balance)[\s:]*(?:rs\.?|₹|inr)?\s*([0-9,]+(?:\.\d{2})?)/gi,
    
    // Currency followed by amount with context words
    /(?:rs\.?|₹|inr)[\s]*([0-9,]+(?:\.\d{2})?)\s*(?:total|final|net|only|due)/gi,
    
    // General total patterns
    /total[\s\S]{0,50}(?:rs\.?|₹|inr)?\s*([0-9,]+(?:\.\d{2})?)/gi,
    
    // Invoice total patterns
    /invoice\s*total[\s:]*(?:rs\.?|₹|inr)?\s*([0-9,]+(?:\.\d{2})?)/gi,
    
    // Simple amount patterns
    /amount[\s:]*(?:rs\.?|₹|inr)?\s*([0-9,]+(?:\.\d{2})?)/gi,
    
    // Any currency amount as fallback
    /(?:rs\.?|₹|inr)[\s]*([0-9,]+(?:\.\d{2})?)/gi
  ];
  
  const amounts = [];
  
  // Find all amounts in the text
  for (const pattern of amountPatterns) {
    const matches = [...text.matchAll(pattern)];
    matches.forEach(match => {
      const amountStr = match[1].replace(/,/g, '');
      const amount = parseFloat(amountStr);
      if (!isNaN(amount) && amount > 0) {
        amounts.push({
          value: amount,
          context: match[0].toLowerCase(),
          priority: getAmountPriority(match[0].toLowerCase())
        });
      }
    });
  }
  
  if (amounts.length > 0) {
    // Sort by priority first, then by value
    amounts.sort((a, b) => {
      if (a.priority !== b.priority) return b.priority - a.priority;
      return b.value - a.value;
    });
    
    const selectedAmount = amounts[0];
    return `Rs ${selectedAmount.value.toLocaleString()}`;
  }
  
  return 'N/A';
}

// Assign priority to amounts based on context words
function getAmountPriority(context) {
  if (context.includes('total') || context.includes('final') || context.includes('net')) return 10;
  if (context.includes('due') || context.includes('payable')) return 8;
  if (context.includes('amount')) return 6;
  if (context.includes('invoice')) return 4;
  return 1; // Default priority for standalone amounts
}

// Extract dates from text in various formats
function extractDateFromText(text) {
  const datePatterns = [
    // DD.MM.YY format (preferred)
    /\b(\d{1,2})\.(\d{1,2})\.(\d{2,4})\b/g,
    // DD/MM/YYYY and DD-MM-YYYY
    /\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\b/g,
    // DD Month YYYY
    /\b(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[.,]?\s+(\d{2,4})\b/gi,
    // Month DD, YYYY
    /\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2})[.,]?\s+(\d{2,4})\b/gi
  ];
  
  for (const pattern of datePatterns) {
    const matches = [...text.matchAll(pattern)];
    if (matches.length > 0) {
      try {
        const match = matches[0];
        let day, month, year;
        
        if (pattern.source.includes('Jan|Feb')) {
          // Handle month name formats
          const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                            'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
          if (match[2]) { // Month DD, YYYY format
            month = monthNames.indexOf(match[1].substring(0, 3)) + 1;
            day = parseInt(match[2]);
            year = parseInt(match[3]);
          } else { // DD Month YYYY format
            day = parseInt(match[1]);
            month = monthNames.indexOf(match[2].substring(0, 3)) + 1;
            year = parseInt(match[3]);
          }
        } else {
          // Handle numeric formats
          day = parseInt(match[1]);
          month = parseInt(match[2]);
          year = parseInt(match[3]);
        }
        
        // Convert 2-digit years to 4-digit
        if (year < 100) {
          year += year > 50 ? 1900 : 2000;
        }
        
        // Validate the date
        const date = new Date(year, month - 1, day);
        if (!isNaN(date.getTime()) && date.getFullYear() === year) {
          return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yy");
        }
      } catch (e) {
        continue; // Try next pattern if this one fails
      }
    }
  }
  return null;
}

// Extract vendor/company names from text
function extractVendorNameFromText(text) {
  const vendorPatterns = [
    // Company names with legal suffixes
    /^([A-Z][a-zA-Z\s&]{2,40}(?:\s+(?:Ltd|Inc|Corp|LLC|Pvt|Private|Limited|LLP|Co)\.?))\s*$/gm,
    
    // Names after "from/vendor/company" keywords
    /(?:from|vendor|company|billed?\s*by|sold\s*by)[\s:]+([A-Z][a-zA-Z\s&]{2,40})/gi,
    
    // Invoice header patterns
    /^([A-Z][a-zA-Z\s&]{3,40})\s*(?:invoice|bill|receipt)/gmi,
    
    // Business names with tax information
    /^([A-Z][a-zA-Z\s&]{3,40})\s*(?:gstin|pan|cin)/gmi,
    
    // Lines starting with proper case names
    /^([A-Z][a-zA-Z\s&]{3,40})\s*$/gm
  ];
  
  for (const pattern of vendorPatterns) {
    const matches = [...text.matchAll(pattern)];
    if (matches.length > 0) {
      let vendor = matches[0][1].trim();
      
      // Clean up the vendor name
      vendor = vendor.replace(/[^\w\s&.-]/g, '').trim();
      vendor = vendor.replace(/\s+/g, ' ');
      
      if (vendor.length >= 3 && vendor.length <= 40) {
        // Skip common non-vendor words
        const excludeWords = ['INVOICE', 'BILL', 'RECEIPT', 'TAX', 'GST', 'TOTAL', 'AMOUNT', 'DATE', 'NUMBER'];
        if (!excludeWords.includes(vendor.toUpperCase())) {
          return vendor.substring(0, 30);
        }
      }
    }
  }
  return null;
}

// Extract invoice/bill numbers from text
function extractInvoiceNumberFromText(text) {
  const invoicePatterns = [
    // Clear invoice/bill number patterns
    /(?:invoice|bill|receipt)[\s#:]*(?:no|number)?[\s#:]*([A-Z0-9\-\/]{3,20})/gi,
    /(?:inv|bill|ref)[\s#:]*(?:no|number)?[\s#:]*([A-Z0-9\-\/]{3,20})/gi,
    
    // Common invoice number formats
    /\b(INV[A-Z0-9\-\/]{2,15})\b/gi,
    /\b([A-Z]{2,4}\d{3,10})\b/g,
    
    // General number patterns
    /(?:no|number|#)[\s:]*([A-Z0-9\-\/]{3,20})/gi,
    
    // Standalone alphanumeric patterns
    /^([A-Z0-9\-\/]{5,15})$/gm
  ];
  
  for (const pattern of invoicePatterns) {
    const matches = [...text.matchAll(pattern)];
    if (matches.length > 0) {
      let invNum = matches[0][1].trim().toUpperCase();
      
      // Validate the format
      if (/^[A-Z0-9\-\/]{3,20}$/.test(invNum)) {
        // Skip false positives
        const excludePatterns = /^(GST|TAX|PAN|TIN|CIN|DATE|AMOUNT|TOTAL)$/;
        if (!excludePatterns.test(invNum)) {
          return invNum;
        }
      }
    }
  }
  return null;
}

// Create or get the main data sheet where results are stored
function getDataSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let dataSheet = spreadsheet.getSheetByName('Data');
  
  if (!dataSheet) {
    dataSheet = spreadsheet.insertSheet('Data');
    const headers = [
      'Timestamp', 
      'Invoice/Bill Date', 
      'Invoice/Bill Number', 
      'Amount', 
      'Vendor/Company Name', 
      'Drive File URL', 
      'File Type'
    ];
    dataSheet.appendRow(headers);
    
    // Format the header row to look nice
    const headerRange = dataSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285F4');
    headerRange.setFontColor('white');
  }
  
  return dataSheet;
}

// Create or get the logs sheet for tracking what the script is doing
function getLogSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  let logSheet = spreadsheet.getSheetByName('Logs');
  
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet('Logs');
    logSheet.appendRow(['Timestamp', 'Message']);
    
    const headerRange = logSheet.getRange(1, 1, 1, 2);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#34A853');
    headerRange.setFontColor('white');
  }
  
  return logSheet;
}

// Write a message to both console and the logs sheet
function logMessage(message, logSheet) {
  console.log(message);
  if (logSheet) {
    try {
      logSheet.appendRow([new Date(), message]);
    } catch (e) {
      console.log(`Failed to log to sheet: ${e.toString()}`);
    }
  }
}

// Find existing folder or create new one in Google Drive
function getOrCreateDriveFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

// Get sender name from email for fallback vendor name
function extractSenderName(email) {
  const sender = email.getFrom();
  const match = sender.match(/^([^<]+)/);
  if (match) {
    return match[0].trim().replace(/[^\w\s]/g, '').substring(0, 30);
  }
  return 'Unknown';
}

// Format date in DD.MM.YY format as required
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd.MM.yy");
}

// Generate a unique invoice number when none is found
function generateInvoiceNumber() {
  return 'INV-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd") + Math.floor(Math.random() * 1000);
}

// Check if file type is supported for processing
function isSupportedFileType(mimeType) {
  const supportedTypes = [
    'image/jpeg', 'image/png', 'image/gif', 'image/bmp', 'image/tiff', 'image/webp',
    'application/pdf',
    'message/rfc822'
  ];
  return supportedTypes.includes(mimeType);
}

// Create Gmail label if it doesn't exist
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

// Get file extension from filename or mime type
function getFileExtension(originalName, mimeType) {
  const nameParts = originalName.split('.');
  if (nameParts.length > 1) return nameParts.pop().toLowerCase();
  
  const extensions = {
    'image/jpeg': 'jpg', 'image/png': 'png', 'image/gif': 'gif', 
    'image/bmp': 'bmp', 'image/tiff': 'tiff', 'image/webp': 'webp',
    'application/pdf': 'pdf', 'message/rfc822': 'eml'
  };
  return extensions[mimeType] || 'dat';
}

// Create filename in required format: Date_Vendor_InvoiceNum_Amount.ext
function formatFilename(data, extension) {
  const cleanVendor = (data.vendorName || 'Unknown').replace(/[^\w\s]/gi, '').replace(/\s+/g, '').substring(0, 15);
  const cleanAmount = (data.totalAmount || 'Rs0').replace(/[^\w]/gi, '');
  const cleanInvoice = (data.invoiceNumber || 'NoNum').replace(/[^\w]/gi, '');
  
  return `${data.invoiceDate}_${cleanVendor}_${cleanInvoice}_${cleanAmount}.${extension}`;
}

// Setup function to initialize everything
function setupScript() {
  const logSheet = getLogSheet();
  logMessage('Setting up Invoice Processing Script', logSheet);
  logMessage('1. Configure DRIVE_FOLDER_ID, SHEET_ID, and OCR_TEMP_FOLDER_ID', logSheet);
  logMessage('2. Enable Drive API in Resources > Advanced Google Services', logSheet);
  logMessage('3. Gemini API is configured for enhanced data extraction', logSheet);
  logMessage('4. Create "docs" folder and upload test documents if needed', logSheet);
  
  getDataSheet();
  getLogSheet();
  
  logMessage('Setup completed successfully', logSheet);
}

// Create automatic trigger to run every 3 hours
function setupTrigger() {
  const logSheet = getLogSheet();
  
  // Remove any existing triggers first
  ScriptApp.getProjectTriggers().forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  
  // Create new trigger for every 3 hours
  ScriptApp.newTrigger('processInvoices')
    .timeBased()
    .everyHours(3)
    .create();
    
  logMessage('Trigger setup completed - will run every 3 hours', logSheet);
}

// Test the main processing function
function testProcessing() {
  const logSheet = getLogSheet();
  logMessage('Testing Invoice Processing with Gemini API', logSheet);
  
  try {
    processInvoices();
    logMessage('Test completed successfully', logSheet);
  } catch (error) {
    logMessage(`Test error: ${error.toString()}`, logSheet);
  }
}

// Test just the Gemini API connection
function testGeminiAPI() {
  const logSheet = getLogSheet();
  logMessage('Testing Gemini API connection', logSheet);
  
  try {
    // Simple test request to Gemini
    const testRequestBody = {
      contents: [{
        parts: [{
          text: "Hello, can you extract invoice data? Please respond with JSON format: {\"test\": \"success\"}"
        }]
      }],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 100
      }
    };

    const response = UrlFetchApp.fetch(`${GEMINI_API_URL}?key=${GEMINI_API_KEY}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(testRequestBody),
      muteHttpExceptions: true
    });

    logMessage(`Gemini API response code: ${response.getResponseCode()}`, logSheet);
    logMessage(`Gemini API response: ${response.getContentText()}`, logSheet);
    
    if (response.getResponseCode() === 200) {
      logMessage('Gemini API connection successful', logSheet);
    } else {
      logMessage('Gemini API connection failed', logSheet);
    }
  } catch (error) {
    logMessage(`Gemini API test error: ${error.toString()}`, logSheet);
  }
}

// Test processing files from docs folder
function testDocsFolder() {
  const logSheet = getLogSheet();
  logMessage('Testing docs folder processing', logSheet);
  
  try {
    const folder = getOrCreateDriveFolder('Viable_Test Documents');
    const tempFolder = DriveApp.getFolderById(OCR_TEMP_FOLDER_ID);
    const dataSheet = getDataSheet();
    
    processDocumentsFromFolder(folder, tempFolder, dataSheet, logSheet);
    logMessage('Docs folder test completed', logSheet);
  } catch (error) {
    logMessage(`Docs folder test error: ${error.toString()}`, logSheet);
  }
}