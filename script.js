// ========================================
// Doc2Any - File Converter JavaScript
// ========================================

// Initialize PDF.js worker
if (typeof pdfjsLib !== 'undefined') {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// DOM Elements
const uploadBox = document.getElementById('uploadBox');
const fileInput = document.getElementById('fileInput');
const conversionModal = document.getElementById('conversionModal');
const closeModal = document.getElementById('closeModal');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const formatButtons = document.getElementById('formatButtons');
const convertBtn = document.getElementById('convertBtn');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const downloadSection = document.getElementById('downloadSection');
const downloadBtn = document.getElementById('downloadBtn');

// State
let selectedFile = null;
let selectedFormat = null;
let convertedBlob = null;
let extractedContent = '';
let preselectedTargetFormat = null; // For format card clicks

// File type icons mapping
const fileIcons = {
    'pdf': 'fa-file-pdf',
    'doc': 'fa-file-word',
    'docx': 'fa-file-word',
    'txt': 'fa-file-lines',
    'odt': 'fa-file-alt',
    'rtf': 'fa-file-contract',
    'xls': 'fa-file-excel',
    'xlsx': 'fa-file-excel',
    'ppt': 'fa-file-powerpoint',
    'pptx': 'fa-file-powerpoint',
    'html': 'fa-file-code',
    'csv': 'fa-file-csv',
    'jpg': 'fa-file-image',
    'jpeg': 'fa-file-image',
    'png': 'fa-file-image',
    'gif': 'fa-file-image',
    'bmp': 'fa-file-image',
    'webp': 'fa-file-image',
    'svg': 'fa-file-image',
    'tiff': 'fa-file-image',
    'tif': 'fa-file-image',
    'ico': 'fa-file-image',
    'default': 'fa-file'
};

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    initializeUpload();
    initializeModal();
    initializeFormatButtons();
    initializeFormatCards(); // Make format cards in the grid clickable
    initializeQuickConvert(); // Quick convert buttons
    initializeAnimations();
    updateStats();
});

// ========================================
// Quick Convert Functionality
// ========================================

function initializeQuickConvert() {
    const quickBtns = document.querySelectorAll('.quick-btn');

    quickBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const format = btn.dataset.format;
            preselectedTargetFormat = format;
            showNotification(`Select a file to convert to ${format.toUpperCase()}`, 'info');
            fileInput.click();
        });
    });
}

// ========================================
// Keyboard Shortcuts
// ========================================

document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + O to open file picker
    if ((e.ctrlKey || e.metaKey) && e.key === 'o') {
        e.preventDefault();
        fileInput.click();
    }
});

// Paste image from clipboard
document.addEventListener('paste', async (e) => {
    const items = e.clipboardData?.items;
    if (!items) return;

    for (const item of items) {
        if (item.type.startsWith('image/')) {
            e.preventDefault();
            const file = item.getAsFile();
            if (file) {
                showNotification('Image pasted from clipboard!', 'success');
                handleFileSelect(file);
            }
            break;
        }
    }
});

// ========================================
// Upload Functionality
// ========================================

function initializeUpload() {
    // Click to upload
    uploadBox.addEventListener('click', () => {
        fileInput.click();
    });

    // File selected
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });

    // Drag and drop
    uploadBox.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadBox.classList.add('dragover');
    });

    uploadBox.addEventListener('dragleave', () => {
        uploadBox.classList.remove('dragover');
    });

    uploadBox.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadBox.classList.remove('dragover');

        if (e.dataTransfer.files.length > 0) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
    });
}

function handleFileSelect(file) {
    // Check file size (100 MB limit)
    const maxSize = 100 * 1024 * 1024;
    if (file.size > maxSize) {
        showNotification('File size exceeds 100 MB limit', 'error');
        return;
    }

    const extension = file.name.split('.').pop().toLowerCase();
    const validExtensions = ['pdf', 'doc', 'docx', 'odt', 'txt', 'rtf', 'xls', 'xlsx', 'ppt', 'pptx', 'html', 'csv', 'jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'svg', 'tiff', 'tif', 'ico'];

    if (!validExtensions.includes(extension)) {
        showNotification('Unsupported file format', 'error');
        return;
    }

    selectedFile = file;
    openModal();
}

// ========================================
// Modal Functionality
// ========================================

function initializeModal() {
    closeModal.addEventListener('click', () => {
        closeModalHandler();
    });

    conversionModal.addEventListener('click', (e) => {
        if (e.target === conversionModal) {
            closeModalHandler();
        }
    });

    // ESC key to close
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && conversionModal.classList.contains('active')) {
            closeModalHandler();
        }
    });

    convertBtn.addEventListener('click', startConversion);
    downloadBtn.addEventListener('click', downloadFile);
}

function openModal() {
    // Update file info
    fileName.textContent = selectedFile.name;
    fileSize.textContent = formatFileSize(selectedFile.size);

    // Update file icon
    const extension = selectedFile.name.split('.').pop().toLowerCase();
    const iconClass = fileIcons[extension] || fileIcons['default'];
    fileInfo.querySelector('.file-icon i').className = `fas ${iconClass}`;

    // Reset modal state
    resetModal();

    // Show available conversion formats based on file type
    updateAvailableFormats(extension);

    // Show modal
    conversionModal.classList.add('active');
    document.body.style.overflow = 'hidden';
}

function closeModalHandler() {
    conversionModal.classList.remove('active');
    document.body.style.overflow = '';
    selectedFile = null;
    selectedFormat = null;
    convertedBlob = null;
    extractedContent = '';
    fileInput.value = '';
}

function resetModal() {
    progressContainer.style.display = 'none';
    downloadSection.style.display = 'none';
    convertBtn.parentElement.style.display = 'block';
    convertBtn.disabled = true;
    progressFill.style.width = '0%';
    progressText.textContent = 'Converting... 0%';

    // Reset format selection
    document.querySelectorAll('.format-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    selectedFormat = null;
}

// All supported output formats for each input type
const conversionMap = {
    'pdf': ['docx', 'txt', 'odt', 'rtf', 'html', 'xlsx', 'pptx', 'csv', 'jpg', 'png'],
    'doc': ['pdf', 'docx', 'txt', 'odt', 'rtf', 'html', 'xlsx', 'pptx', 'csv'],
    'docx': ['pdf', 'txt', 'odt', 'rtf', 'html', 'xlsx', 'pptx', 'csv'],
    'odt': ['pdf', 'docx', 'txt', 'rtf', 'html', 'xlsx', 'pptx', 'csv'],
    'txt': ['pdf', 'docx', 'odt', 'rtf', 'html', 'xlsx', 'pptx', 'csv'],
    'rtf': ['pdf', 'docx', 'txt', 'odt', 'html', 'xlsx', 'pptx', 'csv'],
    'xls': ['pdf', 'xlsx', 'html', 'docx', 'txt', 'csv'],
    'xlsx': ['pdf', 'html', 'docx', 'txt', 'csv'],
    'ppt': ['pdf', 'pptx', 'docx', 'txt', 'csv'],
    'pptx': ['pdf', 'docx', 'txt', 'csv'],
    'html': ['pdf', 'docx', 'txt', 'odt', 'rtf', 'csv'],
    'csv': ['pdf', 'docx', 'txt', 'xlsx', 'html', 'odt', 'rtf'],
    // Image formats
    'jpg': ['png', 'webp', 'gif', 'bmp', 'ico', 'pdf'],
    'jpeg': ['png', 'webp', 'gif', 'bmp', 'ico', 'pdf'],
    'png': ['jpg', 'webp', 'gif', 'bmp', 'ico', 'pdf'],
    'gif': ['jpg', 'png', 'webp', 'bmp', 'ico', 'pdf'],
    'bmp': ['jpg', 'png', 'webp', 'gif', 'ico', 'pdf'],
    'webp': ['jpg', 'png', 'gif', 'bmp', 'ico', 'pdf'],
    'svg': ['jpg', 'png', 'webp', 'gif', 'bmp', 'pdf'],
    'tiff': ['jpg', 'png', 'webp', 'gif', 'bmp', 'pdf'],
    'tif': ['jpg', 'png', 'webp', 'gif', 'bmp', 'pdf'],
    'ico': ['jpg', 'png', 'webp', 'gif', 'bmp', 'pdf']
};

// Image formats list for checking
const imageFormats = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'svg', 'tiff', 'tif', 'ico'];

function updateAvailableFormats(currentExtension) {
    const availableFormats = conversionMap[currentExtension] || ['pdf', 'txt', 'docx', 'html'];

    document.querySelectorAll('.format-btn').forEach(btn => {
        const format = btn.dataset.format;
        if (availableFormats.includes(format)) {
            btn.style.display = 'inline-block';
        } else {
            btn.style.display = 'none';
        }
    });
}

// ========================================
// Format Selection
// ========================================

function initializeFormatButtons() {
    formatButtons.addEventListener('click', (e) => {
        if (e.target.classList.contains('format-btn')) {
            document.querySelectorAll('.format-btn').forEach(btn => {
                btn.classList.remove('active');
            });
            e.target.classList.add('active');
            selectedFormat = e.target.dataset.format;
            convertBtn.disabled = false;
        }
    });
}

// ========================================
// Format Cards (Supported Formats Section)
// ========================================

function initializeFormatCards() {
    const formatsGrid = document.getElementById('formatsGrid');
    if (!formatsGrid) return;

    // Create a hidden file input for format card uploads
    const formatFileInput = document.createElement('input');
    formatFileInput.type = 'file';
    formatFileInput.id = 'formatFileInput';
    formatFileInput.accept = '.doc,.docx,.pdf,.odt,.txt,.rtf,.xls,.xlsx,.ppt,.pptx,.html,.csv,.jpg,.jpeg,.png,.gif,.bmp,.webp,.svg,.tiff,.tif,.ico';
    formatFileInput.style.display = 'none';
    document.body.appendChild(formatFileInput);

    // Handle file selection from format cards
    formatFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelectWithFormat(e.target.files[0]);
        }
        formatFileInput.value = ''; // Reset for next use
    });

    // Add click listeners to format cards
    formatsGrid.addEventListener('click', (e) => {
        const card = e.target.closest('.format-card');
        if (card) {
            const targetFormat = card.dataset.targetFormat;
            if (targetFormat) {
                preselectedTargetFormat = targetFormat;
                showNotification(`Select a file to convert to ${targetFormat.toUpperCase()}`, 'info');
                formatFileInput.click();
            }
        }
    });
}

// Handle file selection with preselected format
function handleFileSelectWithFormat(file) {
    // Check file size (100 MB limit)
    const maxSize = 100 * 1024 * 1024;
    if (file.size > maxSize) {
        showNotification('File size exceeds 100 MB limit', 'error');
        preselectedTargetFormat = null;
        return;
    }

    const extension = file.name.split('.').pop().toLowerCase();
    const validExtensions = ['pdf', 'doc', 'docx', 'odt', 'txt', 'rtf', 'xls', 'xlsx', 'ppt', 'pptx', 'html', 'csv', 'jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'svg', 'tiff', 'tif', 'ico'];

    if (!validExtensions.includes(extension)) {
        showNotification('Unsupported file format', 'error');
        preselectedTargetFormat = null;
        return;
    }

    // Check if the conversion is possible using the global conversionMap
    const availableFormats = conversionMap[extension] || ['pdf', 'txt', 'docx', 'html'];

    // Check if the preselected format is valid for this file type
    if (preselectedTargetFormat && !availableFormats.includes(preselectedTargetFormat)) {
        showNotification(`Cannot convert ${extension.toUpperCase()} to ${preselectedTargetFormat.toUpperCase()}. Please choose a compatible file.`, 'error');
        preselectedTargetFormat = null;
        return;
    }

    selectedFile = file;
    openModalWithFormat();
}

// Open modal with preselected format
function openModalWithFormat() {
    // Update file info
    fileName.textContent = selectedFile.name;
    fileSize.textContent = formatFileSize(selectedFile.size);

    // Update file icon
    const extension = selectedFile.name.split('.').pop().toLowerCase();
    const iconClass = fileIcons[extension] || fileIcons['default'];
    fileInfo.querySelector('.file-icon i').className = `fas ${iconClass}`;

    // Reset modal state
    resetModal();

    // Show available conversion formats based on file type
    updateAvailableFormats(extension);

    // Pre-select the target format if set
    if (preselectedTargetFormat) {
        const targetBtn = document.querySelector(`.format-btn[data-format="${preselectedTargetFormat}"]`);
        if (targetBtn && targetBtn.style.display !== 'none') {
            targetBtn.classList.add('active');
            selectedFormat = preselectedTargetFormat;
            convertBtn.disabled = false;
        }
        preselectedTargetFormat = null; // Reset after use
    }

    // Show modal
    conversionModal.classList.add('active');
    document.body.style.overflow = 'hidden';
}

// ========================================
// Content Extraction Functions
// ========================================

// Extract text from PDF using PDF.js
async function extractTextFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    let fullText = '';

    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map(item => item.str).join(' ');
        fullText += pageText + '\n\n';
    }

    return fullText.trim();
}

// Extract text from DOCX using Mammoth
async function extractTextFromDOCX(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
    return result.value;
}

// Extract HTML from DOCX using Mammoth
async function extractHTMLFromDOCX(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
    return result.value;
}

// Read text file
async function readTextFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = (e) => reject(e);
        reader.readAsText(file);
    });
}

// Extract text from PPTX using JSZip
async function extractTextFromPPTX(file) {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    let fullText = '';
    let slideNumber = 1;

    // PPTX files store slides in ppt/slides/slide1.xml, slide2.xml, etc.
    const slideFiles = Object.keys(zip.files)
        .filter(name => name.match(/ppt\/slides\/slide\d+\.xml$/))
        .sort((a, b) => {
            const numA = parseInt(a.match(/slide(\d+)/)[1]);
            const numB = parseInt(b.match(/slide(\d+)/)[1]);
            return numA - numB;
        });

    for (const slideFile of slideFiles) {
        const content = await zip.files[slideFile].async('string');
        // Extract text from XML - look for <a:t> tags which contain text
        const textMatches = content.match(/<a:t>([^<]*)<\/a:t>/g);
        if (textMatches) {
            const slideText = textMatches
                .map(match => match.replace(/<\/?a:t>/g, ''))
                .filter(text => text.trim())
                .join(' ');
            if (slideText.trim()) {
                fullText += `--- Slide ${slideNumber} ---\n${slideText}\n\n`;
            }
        }
        slideNumber++;
    }

    return fullText.trim() || '[No text content found in presentation]';
}

// Extract text from XLSX using JSZip
async function extractTextFromXLSX(file) {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    let fullText = '';

    // First, get shared strings (XLSX stores text in a shared strings table)
    let sharedStrings = [];
    if (zip.files['xl/sharedStrings.xml']) {
        const ssContent = await zip.files['xl/sharedStrings.xml'].async('string');
        const ssMatches = ssContent.match(/<t[^>]*>([^<]*)<\/t>/g);
        if (ssMatches) {
            sharedStrings = ssMatches.map(match =>
                match.replace(/<\/?t[^>]*>/g, '')
            );
        }
    }

    // Get worksheet data
    const sheetFiles = Object.keys(zip.files)
        .filter(name => name.match(/xl\/worksheets\/sheet\d+\.xml$/))
        .sort();

    let sheetNumber = 1;
    for (const sheetFile of sheetFiles) {
        const content = await zip.files[sheetFile].async('string');
        fullText += `--- Sheet ${sheetNumber} ---\n`;

        // Extract cell values
        const rows = content.match(/<row[^>]*>[\s\S]*?<\/row>/g);
        if (rows) {
            for (const row of rows) {
                const cells = row.match(/<c[^>]*>[\s\S]*?<\/c>/g);
                if (cells) {
                    const rowValues = [];
                    for (const cell of cells) {
                        // Check if it's a shared string reference
                        if (cell.includes('t="s"')) {
                            const vMatch = cell.match(/<v>(\d+)<\/v>/);
                            if (vMatch && sharedStrings[parseInt(vMatch[1])]) {
                                rowValues.push(sharedStrings[parseInt(vMatch[1])]);
                            }
                        } else {
                            // Direct value
                            const vMatch = cell.match(/<v>([^<]*)<\/v>/);
                            if (vMatch) {
                                rowValues.push(vMatch[1]);
                            }
                        }
                    }
                    if (rowValues.length > 0) {
                        fullText += rowValues.join('\t') + '\n';
                    }
                }
            }
        }
        fullText += '\n';
        sheetNumber++;
    }

    return fullText.trim() || '[No data found in spreadsheet]';
}

// Extract text from ODT using JSZip
async function extractTextFromODT(file) {
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    let fullText = '';

    // ODT stores content in content.xml
    if (zip.files['content.xml']) {
        const content = await zip.files['content.xml'].async('string');

        // Extract text from <text:p> and <text:span> tags
        // Remove all XML tags but preserve line breaks for paragraphs
        let text = content
            .replace(/<text:p[^>]*>/g, '\n')
            .replace(/<text:line-break\/>/g, '\n')
            .replace(/<text:tab\/>/g, '\t')
            .replace(/<[^>]+>/g, '')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'");

        fullText = text.trim();
    }

    return fullText || '[No text content found in document]';
}

// Read and parse CSV file
async function readCSVFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const text = e.target.result;
            resolve(text);
        };
        reader.onerror = (e) => reject(e);
        reader.readAsText(file);
    });
}

// Parse CSV to structured data (array of arrays)
function parseCSV(text) {
    const lines = text.split(/\r?\n/);
    const result = [];

    for (const line of lines) {
        if (line.trim()) {
            // Handle quoted fields with commas
            const row = [];
            let field = '';
            let inQuotes = false;

            for (let i = 0; i < line.length; i++) {
                const char = line[i];

                if (char === '"') {
                    if (inQuotes && line[i + 1] === '"') {
                        field += '"';
                        i++;
                    } else {
                        inQuotes = !inQuotes;
                    }
                } else if (char === ',' && !inQuotes) {
                    row.push(field.trim());
                    field = '';
                } else {
                    field += char;
                }
            }
            row.push(field.trim());
            result.push(row);
        }
    }

    return result;
}

// Convert CSV data to formatted text
function csvToText(csvData) {
    return csvData.map(row => row.join('\t')).join('\n');
}

// ========================================
// Conversion Functions
// ========================================

// Convert content to PDF using jsPDF
function convertToPDF(text, originalFileName) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();

    // Set font
    doc.setFont('helvetica');
    doc.setFontSize(12);

    // Add title
    doc.setFontSize(16);
    doc.setFont('helvetica', 'bold');
    doc.text('Converted Document', 105, 20, { align: 'center' });

    doc.setFontSize(10);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(128);
    doc.text(`Original: ${originalFileName}`, 105, 28, { align: 'center' });
    doc.text(`Converted by Doc2Any`, 105, 34, { align: 'center' });

    // Add horizontal line
    doc.setDrawColor(200);
    doc.line(20, 40, 190, 40);

    // Reset text color
    doc.setTextColor(0);
    doc.setFontSize(11);

    // Split text into lines that fit the page
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 20;
    const maxWidth = pageWidth - (margin * 2);
    const lineHeight = 6;
    let yPosition = 50;

    // Split by paragraphs first
    const paragraphs = text.split(/\n\n+/);

    paragraphs.forEach(paragraph => {
        if (paragraph.trim()) {
            const lines = doc.splitTextToSize(paragraph.trim(), maxWidth);

            lines.forEach(line => {
                if (yPosition > pageHeight - margin) {
                    doc.addPage();
                    yPosition = margin;
                }
                doc.text(line, margin, yPosition);
                yPosition += lineHeight;
            });

            yPosition += lineHeight / 2; // Extra space between paragraphs
        }
    });

    return doc.output('blob');
}

// Convert content to HTML
function convertToHTML(text, originalFileName) {
    const escapedText = text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/\n\n+/g, '</p><p>')
        .replace(/\n/g, '<br>');

    const html = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${originalFileName} - Converted by Doc2Any</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
            line-height: 1.6;
            max-width: 800px;
            margin: 0 auto;
            padding: 40px 20px;
            color: #333;
            background: #f9f9f9;
        }
        .header {
            text-align: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 2px solid #e0e0e0;
        }
        .header h1 {
            margin: 0 0 10px 0;
            color: #2c3e50;
        }
        .header p {
            margin: 0;
            color: #7f8c8d;
            font-size: 14px;
        }
        .content {
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .content p {
            margin: 0 0 15px 0;
        }
        .footer {
            text-align: center;
            margin-top: 30px;
            color: #95a5a6;
            font-size: 12px;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>Converted Document</h1>
        <p>Original: ${originalFileName}</p>
    </div>
    <div class="content">
        <p>${escapedText}</p>
    </div>
    <div class="footer">
        Converted by Doc2Any File Converter
    </div>
</body>
</html>`;

    return new Blob([html], { type: 'text/html' });
}

// Convert content to TXT
function convertToTXT(text, originalFileName) {
    const header = `========================================
Document Converted by Doc2Any
Original File: ${originalFileName}
Date: ${new Date().toLocaleString()}
========================================

`;
    return new Blob([header + text], { type: 'text/plain' });
}

// Convert content to DOCX using docx.js
async function convertToDOCX(text, originalFileName) {
    // The docx library exposes itself as 'docx' globally
    let docxLib = window.docx;

    // If not loaded, wait a bit and try again
    if (!docxLib) {
        await new Promise(resolve => setTimeout(resolve, 1000));
        docxLib = window.docx;
    }

    if (!docxLib) {
        throw new Error('DOCX library failed to load. Please check your internet connection and refresh the page.');
    }

    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docxLib;

    // Split text into paragraphs
    const paragraphs = text.split(/\n\n+/).filter(p => p.trim());

    const children = [
        // Title
        new Paragraph({
            children: [
                new TextRun({
                    text: "Converted Document",
                    bold: true,
                    size: 32,
                }),
            ],
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
        }),
        // Subtitle with original filename
        new Paragraph({
            children: [
                new TextRun({
                    text: `Original: ${originalFileName}`,
                    italics: true,
                    color: "666666",
                    size: 20,
                }),
            ],
            alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
            children: [
                new TextRun({
                    text: `Converted by Doc2Any`,
                    color: "666666",
                    size: 18,
                }),
            ],
            alignment: AlignmentType.CENTER,
        }),
        // Empty paragraph for spacing
        new Paragraph({}),
    ];

    // Add content paragraphs
    paragraphs.forEach(para => {
        const lines = para.split('\n');
        const textRuns = [];

        lines.forEach((line, index) => {
            textRuns.push(new TextRun({
                text: line,
                size: 24,
            }));
            if (index < lines.length - 1) {
                textRuns.push(new TextRun({ break: 1 }));
            }
        });

        children.push(new Paragraph({
            children: textRuns,
            spacing: { after: 200 },
        }));
    });

    const doc = new Document({
        sections: [{
            properties: {},
            children: children,
        }],
    });

    const blob = await Packer.toBlob(doc);
    return blob;
}

// Convert content to RTF
function convertToRTF(text, originalFileName) {
    // RTF header and formatting
    const rtfHeader = `{\\rtf1\\ansi\\deff0
{\\fonttbl{\\f0 Arial;}}
{\\colortbl;\\red0\\green0\\blue0;\\red102\\green102\\blue102;}
\\paperw12240\\paperh15840\\margl1440\\margr1440\\margt1440\\margb1440
`;

    // Title
    let rtfContent = `\\pard\\qc\\b\\fs32 Converted Document\\b0\\par
\\pard\\qc\\cf2\\fs20\\i Original: ${escapeRTF(originalFileName)}\\i0\\par
\\pard\\qc\\fs18 Converted by Doc2Any\\cf1\\par
\\par
\\pard\\fs24 `;

    // Content - escape special characters and convert line breaks
    const escapedText = escapeRTF(text);
    rtfContent += escapedText.replace(/\n\n+/g, '\\par\\par ').replace(/\n/g, '\\line ');

    const rtfFooter = `\\par}`;

    return new Blob([rtfHeader + rtfContent + rtfFooter], { type: 'application/rtf' });
}

// Helper function to escape RTF special characters
function escapeRTF(text) {
    return text
        .replace(/\\/g, '\\\\')
        .replace(/\{/g, '\\{')
        .replace(/\}/g, '\\}');
}

// Convert content to ODT (OpenDocument Text)
function convertToODT(text, originalFileName) {
    // ODT is a ZIP file containing XML files
    // For simplicity, we'll create a basic flat ODT (FODT) which is a single XML file
    const escapedText = text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');

    const paragraphs = escapedText.split(/\n\n+/).map(p =>
        `<text:p text:style-name="Standard">${p.replace(/\n/g, '<text:line-break/>')}</text:p>`
    ).join('\n');

    const fodt = `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    office:version="1.2" office:mimetype="application/vnd.oasis.opendocument.text">
    <office:styles>
        <style:style style:name="Title" style:family="paragraph">
            <style:text-properties fo:font-size="18pt" fo:font-weight="bold"/>
        </style:style>
        <style:style style:name="Subtitle" style:family="paragraph">
            <style:text-properties fo:font-size="12pt" fo:color="#666666" fo:font-style="italic"/>
        </style:style>
        <style:style style:name="Standard" style:family="paragraph">
            <style:text-properties fo:font-size="12pt"/>
        </style:style>
    </office:styles>
    <office:body>
        <office:text>
            <text:p text:style-name="Title">Converted Document</text:p>
            <text:p text:style-name="Subtitle">Original: ${originalFileName.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</text:p>
            <text:p text:style-name="Subtitle">Converted by Doc2Any</text:p>
            <text:p text:style-name="Standard"></text:p>
            ${paragraphs}
        </office:text>
    </office:body>
</office:document>`;

    return new Blob([fodt], { type: 'application/vnd.oasis.opendocument.text' });
}

// Convert content to XLSX using ExcelJS
async function convertToXLSX(text, originalFileName) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Doc2Any';
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet('Converted Content');

    // Add title row
    worksheet.mergeCells('A1:D1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = 'Converted Document';
    titleCell.font = { bold: true, size: 16 };
    titleCell.alignment = { horizontal: 'center' };

    // Add metadata row
    worksheet.mergeCells('A2:D2');
    const metaCell = worksheet.getCell('A2');
    metaCell.value = `Original: ${originalFileName} | Converted by Doc2Any`;
    metaCell.font = { italic: true, color: { argb: 'FF666666' }, size: 10 };
    metaCell.alignment = { horizontal: 'center' };

    // Empty row
    worksheet.addRow([]);

    // Split text into lines and add as rows
    const lines = text.split('\n');
    lines.forEach(line => {
        if (line.trim()) {
            worksheet.addRow([line]);
        }
    });

    // Auto-fit column width
    worksheet.getColumn(1).width = 100;

    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

// Convert content to PPTX using PptxGenJS
async function convertToPPTX(text, originalFileName) {
    const pptx = new PptxGenJS();
    pptx.author = 'Doc2Any';
    pptx.title = 'Converted Document';
    pptx.subject = `Converted from ${originalFileName}`;

    // Title slide
    let slide = pptx.addSlide();
    slide.addText('Converted Document', {
        x: 0.5,
        y: 2,
        w: '90%',
        h: 1,
        fontSize: 36,
        bold: true,
        align: 'center',
        color: '363636'
    });
    slide.addText(`Original: ${originalFileName}`, {
        x: 0.5,
        y: 3.2,
        w: '90%',
        h: 0.5,
        fontSize: 14,
        align: 'center',
        color: '666666',
        italic: true
    });
    slide.addText('Converted by Doc2Any', {
        x: 0.5,
        y: 3.7,
        w: '90%',
        h: 0.5,
        fontSize: 12,
        align: 'center',
        color: '999999'
    });

    // Content slides - split text into chunks for each slide
    const paragraphs = text.split(/\n\n+/).filter(p => p.trim());
    const maxCharsPerSlide = 1500;
    let currentText = '';
    let slideNumber = 1;

    for (const para of paragraphs) {
        if ((currentText + para).length > maxCharsPerSlide && currentText) {
            // Create a content slide
            slide = pptx.addSlide();
            slide.addText(currentText, {
                x: 0.5,
                y: 0.5,
                w: '90%',
                h: '85%',
                fontSize: 14,
                color: '363636',
                valign: 'top'
            });
            slide.addText(`Page ${slideNumber}`, {
                x: 0.5,
                y: 5,
                w: '90%',
                h: 0.3,
                fontSize: 10,
                align: 'right',
                color: '999999'
            });
            slideNumber++;
            currentText = para + '\n\n';
        } else {
            currentText += para + '\n\n';
        }
    }

    // Add remaining content
    if (currentText.trim()) {
        slide = pptx.addSlide();
        slide.addText(currentText, {
            x: 0.5,
            y: 0.5,
            w: '90%',
            h: '85%',
            fontSize: 14,
            color: '363636',
            valign: 'top'
        });
        slide.addText(`Page ${slideNumber}`, {
            x: 0.5,
            y: 5,
            w: '90%',
            h: 0.3,
            fontSize: 10,
            align: 'right',
            color: '999999'
        });
    }

    const blob = await pptx.write({ outputType: 'blob' });
    return blob;
}

// Convert content to CSV
function convertToCSV(text, originalFileName) {
    // Split text into lines
    const lines = text.split(/\n/).filter(line => line.trim());

    // Convert each line to CSV format
    const csvLines = lines.map(line => {
        // If line contains tabs, split by tabs (for tabular data)
        if (line.includes('\t')) {
            const cells = line.split('\t');
            return cells.map(cell => {
                // Escape quotes and wrap in quotes if contains comma or quote
                const escaped = cell.replace(/"/g, '""');
                if (escaped.includes(',') || escaped.includes('"') || escaped.includes('\n')) {
                    return `"${escaped}"`;
                }
                return escaped;
            }).join(',');
        }
        // Otherwise treat the whole line as a single cell
        const escaped = line.replace(/"/g, '""');
        if (escaped.includes(',') || escaped.includes('"')) {
            return `"${escaped}"`;
        }
        return escaped;
    });

    const csvContent = `Original File,${originalFileName}\nConverted By,Doc2Any\nDate,${new Date().toLocaleString()}\n\n${csvLines.join('\n')}`;

    return new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
}

// ========================================
// Image Conversion Functions
// ========================================

// Load image from file and return as Image element
function loadImage(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const img = new Image();
            img.onload = () => resolve(img);
            img.onerror = (err) => reject(new Error('Failed to load image'));
            img.src = e.target.result;
        };
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsDataURL(file);
    });
}

// Convert image to specified format using Canvas
async function convertImage(file, targetFormat, quality = 0.92) {
    const img = await loadImage(file);

    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;

    const ctx = canvas.getContext('2d');

    // For formats that don't support transparency, fill with white background
    if (['jpg', 'jpeg', 'bmp'].includes(targetFormat)) {
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
    }

    ctx.drawImage(img, 0, 0);

    // Get mime type for target format
    const mimeTypes = {
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'png': 'image/png',
        'webp': 'image/webp',
        'gif': 'image/gif',
        'bmp': 'image/bmp'
    };

    const mimeType = mimeTypes[targetFormat] || 'image/png';

    return new Promise((resolve, reject) => {
        canvas.toBlob(
            (blob) => {
                if (blob) {
                    resolve(blob);
                } else {
                    reject(new Error('Failed to convert image'));
                }
            },
            mimeType,
            quality
        );
    });
}

// Convert image to ICO format (simplified - creates a PNG wrapped as ICO)
async function convertToICO(file) {
    const img = await loadImage(file);

    // ICO typically uses small sizes, let's use 256x256 max
    const maxSize = 256;
    let width = img.width;
    let height = img.height;

    if (width > maxSize || height > maxSize) {
        if (width > height) {
            height = Math.round((height / width) * maxSize);
            width = maxSize;
        } else {
            width = Math.round((width / height) * maxSize);
            height = maxSize;
        }
    }

    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;

    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0, width, height);

    // Get PNG data
    return new Promise((resolve, reject) => {
        canvas.toBlob(
            (pngBlob) => {
                if (!pngBlob) {
                    reject(new Error('Failed to create ICO'));
                    return;
                }

                // Read the PNG blob
                const reader = new FileReader();
                reader.onload = async () => {
                    const pngData = new Uint8Array(reader.result);

                    // Create ICO file structure
                    // ICO header (6 bytes) + ICO directory entry (16 bytes) + PNG data
                    const icoSize = 6 + 16 + pngData.length;
                    const ico = new Uint8Array(icoSize);

                    // ICO Header
                    ico[0] = 0; ico[1] = 0; // Reserved
                    ico[2] = 1; ico[3] = 0; // Type: 1 = ICO
                    ico[4] = 1; ico[5] = 0; // Number of images: 1

                    // ICO Directory Entry
                    ico[6] = width < 256 ? width : 0; // Width (0 = 256)
                    ico[7] = height < 256 ? height : 0; // Height (0 = 256)
                    ico[8] = 0; // Color palette
                    ico[9] = 0; // Reserved
                    ico[10] = 1; ico[11] = 0; // Color planes
                    ico[12] = 32; ico[13] = 0; // Bits per pixel

                    // Size of image data (4 bytes, little endian)
                    ico[14] = pngData.length & 0xFF;
                    ico[15] = (pngData.length >> 8) & 0xFF;
                    ico[16] = (pngData.length >> 16) & 0xFF;
                    ico[17] = (pngData.length >> 24) & 0xFF;

                    // Offset to image data (4 bytes, little endian) - starts at byte 22
                    ico[18] = 22; ico[19] = 0; ico[20] = 0; ico[21] = 0;

                    // Copy PNG data
                    ico.set(pngData, 22);

                    resolve(new Blob([ico], { type: 'image/x-icon' }));
                };
                reader.onerror = () => reject(new Error('Failed to read PNG data'));
                reader.readAsArrayBuffer(pngBlob);
            },
            'image/png'
        );
    });
}

// Convert image to PDF
async function convertImageToPDF(file) {
    const img = await loadImage(file);
    const { jsPDF } = window.jspdf;

    // Calculate dimensions to fit on page
    const pageWidth = 210; // A4 width in mm
    const pageHeight = 297; // A4 height in mm
    const margin = 10;

    const maxWidth = pageWidth - (margin * 2);
    const maxHeight = pageHeight - (margin * 2);

    let imgWidth = img.width;
    let imgHeight = img.height;

    // Convert pixels to mm (assuming 96 DPI)
    const pxToMm = 0.264583;
    imgWidth *= pxToMm;
    imgHeight *= pxToMm;

    // Scale to fit page
    if (imgWidth > maxWidth) {
        const ratio = maxWidth / imgWidth;
        imgWidth = maxWidth;
        imgHeight *= ratio;
    }
    if (imgHeight > maxHeight) {
        const ratio = maxHeight / imgHeight;
        imgHeight = maxHeight;
        imgWidth *= ratio;
    }

    // Center image on page
    const x = (pageWidth - imgWidth) / 2;
    const y = (pageHeight - imgHeight) / 2;

    // Determine orientation
    const orientation = img.width > img.height ? 'landscape' : 'portrait';
    const doc = new jsPDF(orientation);

    // Get image data URL
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0);
    const imgData = canvas.toDataURL('image/jpeg', 0.95);

    // Add image to PDF
    if (orientation === 'landscape') {
        doc.addImage(imgData, 'JPEG', margin, margin, maxHeight, maxWidth * (img.height / img.width));
    } else {
        doc.addImage(imgData, 'JPEG', x, margin, imgWidth, imgHeight);
    }

    return doc.output('blob');
}

// ========================================
// Image-Preserving Conversion Functions
// ========================================

// Render a single PDF page to a canvas element
async function renderPDFPageToCanvas(pdfDoc, pageNum, scale = 2) {
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;
    return canvas;
}

// Convert PDF pages to a single image (combines pages vertically)
async function convertPDFToImages(file, targetFormat, quality = 0.92) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    const mimeTypes = {
        'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
        'png': 'image/png', 'webp': 'image/webp',
        'gif': 'image/gif', 'bmp': 'image/bmp'
    };
    const mimeType = mimeTypes[targetFormat] || 'image/png';
    const maxPages = Math.min(pdf.numPages, 20);

    const canvases = [];
    let totalHeight = 0;
    let maxWidth = 0;

    for (let i = 1; i <= maxPages; i++) {
        const canvas = await renderPDFPageToCanvas(pdf, i);
        canvases.push(canvas);
        totalHeight += canvas.height;
        maxWidth = Math.max(maxWidth, canvas.width);
    }

    const combinedCanvas = document.createElement('canvas');
    combinedCanvas.width = maxWidth;
    combinedCanvas.height = totalHeight;
    const ctx = combinedCanvas.getContext('2d');

    if (['jpg', 'jpeg', 'bmp'].includes(targetFormat)) {
        ctx.fillStyle = '#FFFFFF';
        ctx.fillRect(0, 0, combinedCanvas.width, combinedCanvas.height);
    }

    let y = 0;
    for (const canvas of canvases) {
        ctx.drawImage(canvas, 0, y);
        y += canvas.height;
    }

    if (pdf.numPages > maxPages) {
        showNotification(`Converted first ${maxPages} of ${pdf.numPages} pages`, 'info');
    }

    return new Promise((resolve, reject) => {
        combinedCanvas.toBlob(
            (blob) => blob ? resolve(blob) : reject(new Error('Failed to create image')),
            mimeType, quality
        );
    });
}

// Convert DOCX to PDF preserving embedded images
async function convertDOCXToPDFWithImages(file, originalFileName) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    const html = result.value;

    if (!html.includes('<img')) {
        return null; // No images found, fall back to text-based conversion
    }

    const parser = new DOMParser();
    const parsedDoc = parser.parseFromString('<div>' + html + '</div>', 'text/html');

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();

    const margin = 20;
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const maxWidth = pageWidth - margin * 2;
    let y = margin;

    function checkPageBreak(needed) {
        if (y + needed > pageHeight - margin) {
            pdf.addPage();
            y = margin;
        }
    }

    async function addImageToPDF(src) {
        if (!src || !src.startsWith('data:')) return;
        try {
            const imgEl = new Image();
            await new Promise((resolve, reject) => {
                imgEl.onload = resolve;
                imgEl.onerror = reject;
                imgEl.src = src;
            });

            const pxToMm = 0.264583;
            let w = imgEl.width * pxToMm;
            let h = imgEl.height * pxToMm;

            if (w > maxWidth) { h *= maxWidth / w; w = maxWidth; }
            const maxH = pageHeight - margin * 2;
            if (h > maxH) { w *= maxH / h; h = maxH; }

            checkPageBreak(h);
            const fmt = src.includes('image/png') ? 'PNG' : 'JPEG';
            pdf.addImage(src, fmt, margin, y, w, h);
            y += h + 5;
        } catch (e) {
            console.warn('Failed to add image to PDF:', e);
        }
    }

    async function processNode(node) {
        if (node.nodeType === Node.TEXT_NODE) {
            const text = node.textContent.trim();
            if (text) {
                const lines = pdf.splitTextToSize(text, maxWidth);
                for (const line of lines) {
                    checkPageBreak(6);
                    pdf.text(line, margin, y);
                    y += 6;
                }
            }
            return;
        }
        if (node.nodeType !== Node.ELEMENT_NODE) return;

        const tag = node.tagName;

        if (tag === 'IMG') {
            await addImageToPDF(node.getAttribute('src'));
            return;
        }

        if (/^H[1-6]$/.test(tag)) {
            const level = parseInt(tag[1]);
            pdf.setFontSize(18 - level * 2);
            pdf.setFont('helvetica', 'bold');
            y += 4;
        }

        for (const child of node.childNodes) {
            await processNode(child);
        }

        if (/^H[1-6]$/.test(tag)) {
            pdf.setFontSize(11);
            pdf.setFont('helvetica', 'normal');
            y += 2;
        }

        if (['P', 'DIV', 'TABLE', 'UL', 'OL', 'LI', 'BLOCKQUOTE'].includes(tag) || /^H[1-6]$/.test(tag)) {
            y += 3;
        }
    }

    const container = parsedDoc.body.firstChild;
    for (const child of container.childNodes) {
        await processNode(child);
    }

    return pdf.output('blob');
}

// Convert PDF to DOCX preserving page visuals as embedded images
async function convertPDFToDOCXWithImages(file, originalFileName) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    const docxLib = window.docx;
    if (!docxLib || !docxLib.ImageRun) return null;

    const { Document, Packer, Paragraph, TextRun, ImageRun, HeadingLevel, AlignmentType } = docxLib;

    const children = [
        new Paragraph({
            children: [new TextRun({ text: "Converted Document", bold: true, size: 32 })],
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
        }),
        new Paragraph({
            children: [new TextRun({ text: 'Original: ' + originalFileName, italics: true, color: "666666", size: 20 })],
            alignment: AlignmentType.CENTER,
        }),
        new Paragraph({}),
    ];

    for (let i = 1; i <= pdf.numPages; i++) {
        const canvas = await renderPDFPageToCanvas(pdf, i, 2);
        const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
        const imgArrayBuffer = await blob.arrayBuffer();

        const targetWidth = 600;
        const targetHeight = Math.round(targetWidth * (canvas.height / canvas.width));

        children.push(
            new Paragraph({
                children: [
                    new ImageRun({
                        data: imgArrayBuffer,
                        transformation: { width: targetWidth, height: targetHeight },
                    }),
                ],
            })
        );

        // Extract text below image for searchability
        const page = await pdf.getPage(i);
        const textContent = await page.getTextContent();
        const pageText = textContent.items.map(item => item.str).join(' ').trim();
        if (pageText) {
            children.push(
                new Paragraph({
                    children: [new TextRun({ text: pageText, size: 20 })],
                    spacing: { after: 400 },
                })
            );
        }
    }

    const doc = new Document({ sections: [{ properties: {}, children }] });
    return await Packer.toBlob(doc);
}

// ========================================
// Main Conversion Process
// ========================================

async function startConversion() {
    if (!selectedFile || !selectedFormat) return;

    // Show progress
    convertBtn.parentElement.style.display = 'none';
    progressContainer.style.display = 'block';

    // Progress animation
    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.random() * 10;
        if (progress > 90) progress = 90;
        progressFill.style.width = `${progress}%`;
        progressText.textContent = `Converting... ${Math.round(progress)}%`;
    }, 150);

    try {
        const extension = selectedFile.name.split('.').pop().toLowerCase();
        const originalFileName = selectedFile.name;

        // Check if this is an image file
        const isImageFile = imageFormats.includes(extension);
        const isImageOutput = ['jpg', 'jpeg', 'png', 'webp', 'gif', 'bmp', 'ico'].includes(selectedFormat);

        // Handle image-to-image or image-to-PDF conversion
        if (isImageFile) {
            progressText.textContent = 'Converting image...';

            if (selectedFormat === 'pdf') {
                convertedBlob = await convertImageToPDF(selectedFile);
            } else if (selectedFormat === 'ico') {
                convertedBlob = await convertToICO(selectedFile);
            } else if (isImageOutput) {
                convertedBlob = await convertImage(selectedFile, selectedFormat);
            } else {
                throw new Error(`Cannot convert image to ${selectedFormat.toUpperCase()}`);
            }

            clearInterval(interval);
            progressFill.style.width = '100%';
            progressText.textContent = 'Conversion complete!';

            setTimeout(() => {
                progressContainer.style.display = 'none';
                downloadSection.style.display = 'block';
                incrementConversionCount(selectedFile ? selectedFile.size : 0);
            }, 500);

            return;
        }

        // ========================================
        // Image-Preserving Conversion Paths
        // ========================================

        // PDF → Image formats: render pages visually instead of extracting text
        if (extension === 'pdf' && isImageOutput) {
            progressText.textContent = 'Rendering PDF pages...';
            if (selectedFormat === 'ico') {
                const pdfArrayBuffer = await selectedFile.arrayBuffer();
                const pdfDoc = await pdfjsLib.getDocument({ data: pdfArrayBuffer }).promise;
                const canvas = await renderPDFPageToCanvas(pdfDoc, 1);
                const pngBlob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
                const tempFile = new File([pngBlob], 'page.png', { type: 'image/png' });
                convertedBlob = await convertToICO(tempFile);
            } else {
                convertedBlob = await convertPDFToImages(selectedFile, selectedFormat);
            }

            clearInterval(interval);
            progressFill.style.width = '100%';
            progressText.textContent = 'Conversion complete!';
            setTimeout(() => {
                progressContainer.style.display = 'none';
                downloadSection.style.display = 'block';
                incrementConversionCount(selectedFile ? selectedFile.size : 0);
            }, 500);
            return;
        }

        // PDF → DOCX: render pages as images and embed them in the document
        if (extension === 'pdf' && selectedFormat === 'docx') {
            progressText.textContent = 'Converting PDF with images...';
            try {
                convertedBlob = await convertPDFToDOCXWithImages(selectedFile, originalFileName);
            } catch (e) {
                console.warn('Image-preserving PDF→DOCX failed, falling back to text:', e);
                convertedBlob = null;
            }
            if (convertedBlob) {
                clearInterval(interval);
                progressFill.style.width = '100%';
                progressText.textContent = 'Conversion complete!';
                setTimeout(() => {
                    progressContainer.style.display = 'none';
                    downloadSection.style.display = 'block';
                    incrementConversionCount(selectedFile ? selectedFile.size : 0);
                }, 500);
                return;
            }
        }

        // DOCX → PDF: preserve embedded images from the document
        if (extension === 'docx' && selectedFormat === 'pdf') {
            progressText.textContent = 'Converting with images...';
            try {
                convertedBlob = await convertDOCXToPDFWithImages(selectedFile, originalFileName);
            } catch (e) {
                console.warn('Image-preserving DOCX→PDF failed, falling back to text:', e);
                convertedBlob = null;
            }
            if (convertedBlob) {
                clearInterval(interval);
                progressFill.style.width = '100%';
                progressText.textContent = 'Conversion complete!';
                setTimeout(() => {
                    progressContainer.style.display = 'none';
                    downloadSection.style.display = 'block';
                    incrementConversionCount(selectedFile ? selectedFile.size : 0);
                }, 500);
                return;
            }
        }

        // Step 1: Extract content from source file
        progressText.textContent = 'Reading file...';

        let content = '';

        if (extension === 'pdf') {
            if (typeof pdfjsLib === 'undefined') {
                throw new Error('PDF.js library not loaded');
            }
            content = await extractTextFromPDF(selectedFile);
        } else if (extension === 'docx') {
            if (typeof mammoth === 'undefined') {
                throw new Error('Mammoth library not loaded');
            }
            if (selectedFormat === 'html') {
                content = await extractHTMLFromDOCX(selectedFile);
            } else {
                content = await extractTextFromDOCX(selectedFile);
            }
        } else if (extension === 'pptx' || extension === 'ppt') {
            if (typeof JSZip === 'undefined') {
                throw new Error('JSZip library not loaded');
            }
            content = await extractTextFromPPTX(selectedFile);
        } else if (extension === 'xlsx' || extension === 'xls') {
            if (typeof JSZip === 'undefined') {
                throw new Error('JSZip library not loaded');
            }
            content = await extractTextFromXLSX(selectedFile);
        } else if (extension === 'txt' || extension === 'html' || extension === 'rtf') {
            content = await readTextFile(selectedFile);
            // Strip HTML tags if source is HTML and converting to other format
            if (extension === 'html' && selectedFormat !== 'html') {
                content = content.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
            }
        } else if (extension === 'odt') {
            // ODT files are also ZIP-based XML
            if (typeof JSZip === 'undefined') {
                throw new Error('JSZip library not loaded');
            }
            content = await extractTextFromODT(selectedFile);
        } else if (extension === 'csv') {
            // Read CSV file as text
            const csvText = await readCSVFile(selectedFile);
            const csvData = parseCSV(csvText);
            content = csvToText(csvData);
        } else {
            // For unsupported extraction, use text reader as fallback
            try {
                content = await readTextFile(selectedFile);
            } catch {
                content = `[Content from ${originalFileName}]\n\nThis file format requires server-side processing for full conversion.\nBasic text extraction was attempted.`;
            }
        }

        if (!content || content.trim().length === 0) {
            content = `[No text content could be extracted from ${originalFileName}]\n\nThis may be because:\n- The file is image-based (scanned document)\n- The file is encrypted or protected\n- The file format is not fully supported for text extraction`;
        }

        // Step 2: Convert to target format
        progressText.textContent = 'Converting...';

        switch (selectedFormat) {
            case 'pdf':
                convertedBlob = convertToPDF(content, originalFileName);
                break;
            case 'html':
                if (extension === 'docx' && content.includes('<')) {
                    // Already HTML from mammoth
                    const wrappedHTML = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${originalFileName} - Converted by Doc2Any</title>
    <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; max-width: 800px; margin: 0 auto; padding: 40px 20px; }
        img { max-width: 100%; }
    </style>
</head>
<body>
${content}
<footer style="margin-top: 40px; text-align: center; color: #888; font-size: 12px;">
    Converted by Doc2Any File Converter
</footer>
</body>
</html>`;
                    convertedBlob = new Blob([wrappedHTML], { type: 'text/html' });
                } else {
                    convertedBlob = convertToHTML(content, originalFileName);
                }
                break;
            case 'txt':
                convertedBlob = convertToTXT(content, originalFileName);
                break;
            case 'docx':
                convertedBlob = await convertToDOCX(content, originalFileName);
                break;
            case 'odt':
                convertedBlob = convertToODT(content, originalFileName);
                break;
            case 'rtf':
                convertedBlob = convertToRTF(content, originalFileName);
                break;
            case 'xlsx':
                convertedBlob = await convertToXLSX(content, originalFileName);
                break;
            case 'pptx':
                convertedBlob = await convertToPPTX(content, originalFileName);
                break;
            case 'csv':
                convertedBlob = convertToCSV(content, originalFileName);
                break;
            default:
                throw new Error(`Unsupported output format: ${selectedFormat}`);
        }

        clearInterval(interval);

        // Complete progress
        progressFill.style.width = '100%';
        progressText.textContent = 'Conversion complete!';

        // Show download section
        setTimeout(() => {
            progressContainer.style.display = 'none';
            downloadSection.style.display = 'block';
            incrementConversionCount(selectedFile ? selectedFile.size : 0);
        }, 500);

    } catch (error) {
        clearInterval(interval);
        console.error('Conversion error:', error);
        showNotification(`Conversion failed: ${error.message}`, 'error');
        resetModal();
        convertBtn.parentElement.style.display = 'block';
    }
}

function downloadFile() {
    if (!convertedBlob) return;

    const originalName = selectedFile.name.split('.').slice(0, -1).join('.');
    // Use .fodt extension for flat ODT files
    const extension = selectedFormat === 'odt' ? 'fodt' : selectedFormat;
    const newFileName = `${originalName}_converted.${extension}`;

    const url = URL.createObjectURL(convertedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = newFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showNotification('File downloaded successfully!', 'success');

    // Close modal after download
    setTimeout(() => {
        closeModalHandler();
    }, 1000);
}

// ========================================
// Utility Functions
// ========================================

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';

    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function showNotification(message, type = 'info') {
    // Remove existing notifications
    const existing = document.querySelector('.notification');
    if (existing) existing.remove();

    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <i class="fas ${type === 'success' ? 'fa-check-circle' : type === 'error' ? 'fa-exclamation-circle' : 'fa-info-circle'}"></i>
        <span>${message}</span>
        <button class="notification-close" onclick="this.parentElement.remove()">
            <i class="fas fa-times"></i>
        </button>
    `;

    // Add notification styles dynamically
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 15px 45px 15px 20px;
        border-radius: 12px;
        background: ${type === 'success' ? '#68d391' : type === 'error' ? '#fc8181' : '#5B6EE1'};
        color: white;
        font-weight: 500;
        display: flex;
        align-items: center;
        gap: 12px;
        z-index: 3000;
        animation: slideIn 0.3s ease;
        box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        max-width: 400px;
    `;

    // Style the close button
    const closeBtn = notification.querySelector('.notification-close');
    if (closeBtn) {
        closeBtn.style.cssText = `
            position: absolute;
            right: 10px;
            top: 50%;
            transform: translateY(-50%);
            background: rgba(255,255,255,0.2);
            border: none;
            color: white;
            width: 24px;
            height: 24px;
            border-radius: 50%;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
        `;
    }

    // Add animation keyframes
    if (!document.querySelector('#notification-styles')) {
        const style = document.createElement('style');
        style.id = 'notification-styles';
        style.textContent = `
            @keyframes slideIn {
                from { transform: translateX(100%); opacity: 0; }
                to { transform: translateX(0); opacity: 1; }
            }
            @keyframes slideOut {
                from { transform: translateX(0); opacity: 1; }
                to { transform: translateX(100%); opacity: 0; }
            }
        `;
        document.head.appendChild(style);
    }

    document.body.appendChild(notification);

    // Auto remove
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease forwards';
        setTimeout(() => notification.remove(), 300);
    }, 4000);
}

// ========================================
// Stats - Track User's Real Conversions
// ========================================

function initializeStats() {
    // Get stored values or start from 0
    let fileCount = localStorage.getItem('doc2any_conversions') || 0;
    let totalSizeBytes = localStorage.getItem('doc2any_total_size') || 0;

    fileCount = parseInt(fileCount);
    totalSizeBytes = parseFloat(totalSizeBytes);

    // Reset if values are unrealistically high (legacy dummy data)
    if (fileCount > 1000000 || totalSizeBytes > 1000000000000) {
        fileCount = 0;
        totalSizeBytes = 0;
        localStorage.setItem('doc2any_conversions', 0);
        localStorage.setItem('doc2any_total_size', 0);
    }

    // Update display
    updateStatsDisplay(fileCount, totalSizeBytes);
}

function updateStatsDisplay(fileCount, totalSizeBytes) {
    const totalFilesEl = document.getElementById('totalFiles');
    const totalSizeEl = document.getElementById('totalSize');

    if (totalFilesEl) {
        totalFilesEl.textContent = parseInt(fileCount).toLocaleString();
    }
    if (totalSizeEl) {
        // Display in appropriate unit (Bytes, KB, MB, GB)
        totalSizeEl.textContent = formatSizeForStats(totalSizeBytes);
    }
}

function formatSizeForStats(bytes) {
    if (bytes === 0) return '0';

    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    if (i === 0) return bytes + ' Bytes';
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)).toLocaleString();
}

function getSizeUnit(bytes) {
    if (bytes === 0) return 'MB';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return sizes[i];
}

function incrementConversionCount(fileSizeBytes = 0) {
    // Update file count
    let fileCount = localStorage.getItem('doc2any_conversions') || 0;
    fileCount = parseInt(fileCount) + 1;
    localStorage.setItem('doc2any_conversions', fileCount);

    // Update total size
    let totalSizeBytes = localStorage.getItem('doc2any_total_size') || 0;
    totalSizeBytes = parseFloat(totalSizeBytes) + fileSizeBytes;
    localStorage.setItem('doc2any_total_size', totalSizeBytes);

    // Update display with animation
    const totalFilesEl = document.getElementById('totalFiles');
    const totalSizeEl = document.getElementById('totalSize');

    if (totalFilesEl) {
        totalFilesEl.textContent = fileCount.toLocaleString();
        // Add pulse animation
        totalFilesEl.style.transform = 'scale(1.2)';
        totalFilesEl.style.transition = 'transform 0.3s';
        setTimeout(() => {
            totalFilesEl.style.transform = 'scale(1)';
        }, 300);
    }

    if (totalSizeEl) {
        totalSizeEl.textContent = formatSizeForStats(totalSizeBytes);
        // Update the unit label in the HTML
        updateSizeUnit(totalSizeBytes);
    }
}

function updateSizeUnit(bytes) {
    const statsText = document.querySelector('.hero-stats');
    if (statsText) {
        const unit = getSizeUnit(bytes);
        // Update the text to show correct unit
        const html = statsText.innerHTML;
        statsText.innerHTML = html.replace(/(Bytes|KB|MB|GB)(?=\s*$)/, unit);
    }
}

// Legacy function for compatibility
function updateStats() {
    initializeStats();
}

function animateNumber(element, target) {
    const duration = 2000;
    const start = target - 1000;
    const startTime = performance.now();

    function update(currentTime) {
        const elapsed = currentTime - startTime;
        const progress = Math.min(elapsed / duration, 1);

        const easeOut = 1 - Math.pow(1 - progress, 3);
        const current = Math.floor(start + (target - start) * easeOut);

        element.textContent = current.toLocaleString();

        if (progress < 1) {
            requestAnimationFrame(update);
        }
    }

    requestAnimationFrame(update);
}

// ========================================
// Animations
// ========================================

function initializeAnimations() {
    const chartObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                animateChart();
                chartObserver.unobserve(entry.target);
            }
        });
    }, { threshold: 0.5 });

    const chart = document.querySelector('.donut-chart');
    if (chart) {
        chartObserver.observe(chart);
    }

    const fadeObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.opacity = '1';
                entry.target.style.transform = 'translateY(0)';
            }
        });
    }, { threshold: 0.1 });

    document.querySelectorAll('.step-card, .format-card, .feature-text, .stats-card').forEach(el => {
        el.style.opacity = '0';
        el.style.transform = 'translateY(30px)';
        el.style.transition = 'opacity 0.6s ease, transform 0.6s ease';
        fadeObserver.observe(el);
    });
}

function animateChart() {
    const segments = document.querySelectorAll('.chart-segment');
    segments.forEach((segment, index) => {
        segment.style.transition = 'stroke-dasharray 1s ease';
        segment.style.transitionDelay = `${index * 0.2}s`;
    });

    const chartValue = document.querySelector('.chart-value');
    if (chartValue) {
        let start = 0;
        const end = 1258;
        const duration = 1500;
        const startTime = performance.now();

        function updateValue(currentTime) {
            const elapsed = currentTime - startTime;
            const progress = Math.min(elapsed / duration, 1);
            const current = Math.floor(start + (end - start) * progress);
            chartValue.textContent = current.toLocaleString();

            if (progress < 1) {
                requestAnimationFrame(updateValue);
            }
        }

        requestAnimationFrame(updateValue);
    }
}

// ========================================
// Mobile Menu
// ========================================

// Mobile menu is now handled by initializeMobileMenu()

// ========================================
// Smooth Scrolling
// ========================================

document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const href = this.getAttribute('href');
        if (href === '#') return;

        const target = document.querySelector(href);
        if (target) {
            // Update active nav link
            document.querySelectorAll('.nav-link').forEach(link => {
                link.classList.remove('active');
            });
            this.classList.add('active');

            // Scroll to target with offset for header
            const headerHeight = document.querySelector('.header').offsetHeight;
            const targetPosition = target.getBoundingClientRect().top + window.pageYOffset - headerHeight;

            window.scrollTo({
                top: targetPosition,
                behavior: 'smooth'
            });
        }
    });
});

// Update active nav link on scroll
window.addEventListener('scroll', () => {
    const sections = ['convert', 'ocr', 'help'];
    const headerHeight = document.querySelector('.header').offsetHeight;

    let current = '';
    sections.forEach(id => {
        const section = document.getElementById(id);
        if (section) {
            const sectionTop = section.offsetTop - headerHeight - 100;
            if (window.pageYOffset >= sectionTop) {
                current = id;
            }
        }
    });

    document.querySelectorAll('.nav-link').forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('href') === `#${current}`) {
            link.classList.add('active');
        }
    });
});

// ========================================
// FAQ Toggle
// ========================================

function toggleFaq(element) {
    const faqItem = element.parentElement;
    const isActive = faqItem.classList.contains('active');

    // Close all other FAQs
    document.querySelectorAll('.faq-item').forEach(item => {
        item.classList.remove('active');
    });

    // Toggle current FAQ
    if (!isActive) {
        faqItem.classList.add('active');
    }
}

// ========================================
// Copy Code Button
// ========================================

function copyCode() {
    const codeBlock = document.querySelector('.api-code-block code');
    if (codeBlock) {
        navigator.clipboard.writeText(codeBlock.textContent).then(() => {
            showNotification('Code copied to clipboard!', 'success');
        }).catch(() => {
            showNotification('Failed to copy code', 'error');
        });
    }
}

// ========================================
// OCR Functionality with Tesseract.js
// ========================================

// OCR State
let currentOCRFile = null;
let ocrExtractedText = '';
let ocrConfidence = 0;

function initializeOCR() {
    const ocrUploadBox = document.getElementById('ocrUploadBox');
    const ocrFileInput = document.getElementById('ocrFileInput');
    const ocrModal = document.getElementById('ocrModal');
    const closeOcrModal = document.getElementById('closeOcrModal');
    const copyOcrText = document.getElementById('copyOcrText');
    const downloadOcrText = document.getElementById('downloadOcrText');
    const convertOcrResult = document.getElementById('convertOcrResult');
    const newOcrScan = document.getElementById('newOcrScan');

    if (ocrUploadBox && ocrFileInput) {
        // Click to upload
        ocrUploadBox.addEventListener('click', () => {
            ocrFileInput.click();
        });

        // Drag over
        ocrUploadBox.addEventListener('dragover', (e) => {
            e.preventDefault();
            ocrUploadBox.style.borderColor = '#667eea';
            ocrUploadBox.style.background = 'rgba(255, 255, 255, 1)';
        });

        // Drag leave
        ocrUploadBox.addEventListener('dragleave', () => {
            ocrUploadBox.style.borderColor = 'rgba(102, 126, 234, 0.5)';
            ocrUploadBox.style.background = 'rgba(255, 255, 255, 0.95)';
        });

        // Drop
        ocrUploadBox.addEventListener('drop', (e) => {
            e.preventDefault();
            ocrUploadBox.style.borderColor = 'rgba(102, 126, 234, 0.5)';
            ocrUploadBox.style.background = 'rgba(255, 255, 255, 0.95)';

            if (e.dataTransfer.files.length > 0) {
                handleOCRFile(e.dataTransfer.files[0]);
            }
        });

        // File input change
        ocrFileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleOCRFile(e.target.files[0]);
            }
        });
    }

    // Modal controls
    if (closeOcrModal) {
        closeOcrModal.addEventListener('click', closeOCRModal);
    }

    if (ocrModal) {
        ocrModal.addEventListener('click', (e) => {
            if (e.target === ocrModal) {
                closeOCRModal();
            }
        });
    }

    // Copy OCR text
    if (copyOcrText) {
        copyOcrText.addEventListener('click', () => {
            const textArea = document.getElementById('ocrResultText');
            if (textArea && textArea.value) {
                navigator.clipboard.writeText(textArea.value).then(() => {
                    showNotification('Text copied to clipboard!', 'success');
                }).catch(() => {
                    // Fallback for older browsers
                    textArea.select();
                    document.execCommand('copy');
                    showNotification('Text copied to clipboard!', 'success');
                });
            }
        });
    }

    // Download OCR text
    if (downloadOcrText) {
        downloadOcrText.addEventListener('click', () => {
            const textArea = document.getElementById('ocrResultText');
            if (textArea && textArea.value) {
                const blob = new Blob([textArea.value], { type: 'text/plain' });
                const fileName = currentOCRFile ?
                    currentOCRFile.name.replace(/\.[^/.]+$/, '') + '_ocr.txt' :
                    'ocr_result.txt';
                saveAs(blob, fileName);
                showNotification('Text file downloaded!', 'success');
            }
        });
    }

    // Convert OCR result to document
    if (convertOcrResult) {
        convertOcrResult.addEventListener('click', () => {
            const textArea = document.getElementById('ocrResultText');
            if (textArea && textArea.value) {
                // Create a virtual file from OCR text and open conversion modal
                const ocrText = textArea.value;
                const fileName = currentOCRFile ?
                    currentOCRFile.name.replace(/\.[^/.]+$/, '') + '_ocr.txt' :
                    'ocr_result.txt';

                // Create blob and file
                const blob = new Blob([ocrText], { type: 'text/plain' });
                const file = new File([blob], fileName, { type: 'text/plain' });

                // Close OCR modal and open conversion modal
                closeOCRModal();
                selectedFile = file;
                openModal();
                showNotification('Select a format to convert the OCR text', 'info');
            }
        });
    }

    // New OCR scan
    if (newOcrScan) {
        newOcrScan.addEventListener('click', () => {
            closeOCRModal();
            // Scroll to OCR section
            const ocrSection = document.getElementById('ocr');
            if (ocrSection) {
                const headerHeight = document.querySelector('.header').offsetHeight;
                const targetPosition = ocrSection.getBoundingClientRect().top + window.pageYOffset - headerHeight;
                window.scrollTo({
                    top: targetPosition,
                    behavior: 'smooth'
                });
            }
            // Trigger file input
            setTimeout(() => {
                const ocrFileInput = document.getElementById('ocrFileInput');
                if (ocrFileInput) {
                    ocrFileInput.click();
                }
            }, 500);
        });
    }

    // ESC key to close OCR modal
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            const ocrModal = document.getElementById('ocrModal');
            if (ocrModal && ocrModal.classList.contains('active')) {
                closeOCRModal();
            }
        }
    });
}

async function handleOCRFile(file) {
    // Validate file type
    const validTypes = ['image/jpeg', 'image/png', 'image/tiff', 'image/bmp', 'image/webp'];
    const validExtensions = ['.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp', '.webp'];

    const extension = '.' + file.name.split('.').pop().toLowerCase();

    if (!validTypes.includes(file.type) && !validExtensions.includes(extension)) {
        showNotification('Please upload a valid image file (JPG, PNG, TIFF, BMP, WEBP)', 'error');
        return;
    }

    // Check file size (max 10MB for OCR)
    if (file.size > 10 * 1024 * 1024) {
        showNotification('Image size should be less than 10 MB for OCR', 'error');
        return;
    }

    currentOCRFile = file;

    // Get selected language
    const languageSelect = document.getElementById('ocrLanguage');
    const language = languageSelect ? languageSelect.value : 'eng';

    // Show processing state
    const ocrUploadBox = document.getElementById('ocrUploadBox');
    const ocrProgress = document.getElementById('ocrProgress');
    const ocrProgressFill = document.getElementById('ocrProgressFill');
    const ocrProgressText = document.getElementById('ocrProgressText');
    const ocrUploadIcon = document.getElementById('ocrUploadIcon');
    const ocrUploadText = document.getElementById('ocrUploadText');
    const ocrUploadHint = document.getElementById('ocrUploadHint');

    if (ocrUploadBox) {
        ocrUploadBox.classList.add('processing');
    }

    if (ocrProgress) {
        ocrProgress.style.display = 'block';
    }

    if (ocrUploadIcon) {
        ocrUploadIcon.className = 'fas fa-spinner fa-spin';
    }

    if (ocrUploadText) {
        ocrUploadText.textContent = 'Processing image...';
    }

    if (ocrUploadHint) {
        ocrUploadHint.textContent = 'Please wait while we extract text';
    }

    try {
        // Check if Tesseract is loaded
        if (typeof Tesseract === 'undefined') {
            throw new Error('OCR library not loaded. Please refresh the page.');
        }

        // Perform OCR with Tesseract.js
        const result = await Tesseract.recognize(file, language, {
            logger: (m) => {
                if (m.status === 'recognizing text') {
                    const progress = Math.round(m.progress * 100);
                    if (ocrProgressFill) {
                        ocrProgressFill.style.width = `${progress}%`;
                    }
                    if (ocrProgressText) {
                        ocrProgressText.textContent = `Recognizing text... ${progress}%`;
                    }
                } else if (m.status === 'loading language traineddata') {
                    if (ocrProgressText) {
                        ocrProgressText.textContent = 'Loading language data...';
                    }
                } else if (m.status === 'initializing api') {
                    if (ocrProgressText) {
                        ocrProgressText.textContent = 'Initializing OCR engine...';
                    }
                }
            }
        });

        // Store results
        ocrExtractedText = result.data.text;
        ocrConfidence = result.data.confidence;

        // Reset upload box
        resetOCRUploadBox();

        // Show results in modal
        showOCRResults(file, ocrExtractedText, ocrConfidence);

    } catch (error) {
        console.error('OCR Error:', error);
        showNotification(`OCR failed: ${error.message}`, 'error');
        resetOCRUploadBox();
    }
}

function resetOCRUploadBox() {
    const ocrUploadBox = document.getElementById('ocrUploadBox');
    const ocrProgress = document.getElementById('ocrProgress');
    const ocrProgressFill = document.getElementById('ocrProgressFill');
    const ocrUploadIcon = document.getElementById('ocrUploadIcon');
    const ocrUploadText = document.getElementById('ocrUploadText');
    const ocrUploadHint = document.getElementById('ocrUploadHint');
    const ocrFileInput = document.getElementById('ocrFileInput');

    if (ocrUploadBox) {
        ocrUploadBox.classList.remove('processing');
    }

    if (ocrProgress) {
        ocrProgress.style.display = 'none';
    }

    if (ocrProgressFill) {
        ocrProgressFill.style.width = '0%';
    }

    if (ocrUploadIcon) {
        ocrUploadIcon.className = 'fas fa-camera';
    }

    if (ocrUploadText) {
        ocrUploadText.textContent = 'Drop image or scanned document here';
    }

    if (ocrUploadHint) {
        ocrUploadHint.textContent = 'Supports JPG, PNG, TIFF, BMP, WEBP';
    }

    if (ocrFileInput) {
        ocrFileInput.value = '';
    }
}

function showOCRResults(file, text, confidence) {
    const ocrModal = document.getElementById('ocrModal');
    const ocrPreviewImage = document.getElementById('ocrPreviewImage');
    const ocrResultText = document.getElementById('ocrResultText');
    const ocrCharCount = document.getElementById('ocrCharCount');
    const ocrWordCount = document.getElementById('ocrWordCount');
    const ocrConfidence = document.getElementById('ocrConfidence');

    // Create image preview
    if (ocrPreviewImage && file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            ocrPreviewImage.src = e.target.result;
        };
        reader.readAsDataURL(file);
    }

    // Set extracted text
    if (ocrResultText) {
        ocrResultText.value = text || '[No text detected in the image]';
        ocrResultText.readOnly = false; // Allow editing
    }

    // Calculate stats
    const charCount = text ? text.length : 0;
    const wordCount = text ? text.trim().split(/\s+/).filter(w => w).length : 0;

    if (ocrCharCount) {
        ocrCharCount.textContent = `${charCount.toLocaleString()} characters`;
    }

    if (ocrWordCount) {
        ocrWordCount.textContent = `${wordCount.toLocaleString()} words`;
    }

    if (ocrConfidence) {
        ocrConfidence.textContent = `Confidence: ${Math.round(confidence)}%`;
    }

    // Show modal
    if (ocrModal) {
        ocrModal.classList.add('active');
        document.body.style.overflow = 'hidden';
    }

    showNotification('Text extracted successfully!', 'success');
}

function closeOCRModal() {
    const ocrModal = document.getElementById('ocrModal');

    if (ocrModal) {
        ocrModal.classList.remove('active');
        document.body.style.overflow = '';
    }

    // Reset state
    currentOCRFile = null;
    ocrExtractedText = '';
    ocrConfidence = 0;
}

// Initialize OCR on page load
document.addEventListener('DOMContentLoaded', () => {
    initializeOCR();
    initializeDarkMode();
    initializeScrollProgress();
    initializeBackToTop();
    initializeTypingAnimation();
    initializeMobileMenu();
    initializeRippleEffect();
});

// ========================================
// Dark Mode
// ========================================

function initializeDarkMode() {
    const themeToggle = document.getElementById('themeToggle');
    const themeToggleMobile = document.getElementById('themeToggleMobile');
    const saved = localStorage.getItem('doc2any_theme');

    // Apply saved theme or detect system preference
    if (saved === 'dark' || (!saved && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
        document.documentElement.setAttribute('data-theme', 'dark');
        updateThemeIcons(true);
    }

    function toggle() {
        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        if (isDark) {
            document.documentElement.removeAttribute('data-theme');
            localStorage.setItem('doc2any_theme', 'light');
        } else {
            document.documentElement.setAttribute('data-theme', 'dark');
            localStorage.setItem('doc2any_theme', 'dark');
        }
        updateThemeIcons(!isDark);
    }

    function updateThemeIcons(isDark) {
        const icon = isDark ? 'fa-sun' : 'fa-moon';
        if (themeToggle) themeToggle.querySelector('i').className = 'fas ' + icon;
        if (themeToggleMobile) {
            themeToggleMobile.querySelector('i').className = 'fas ' + icon;
            themeToggleMobile.querySelector('span').textContent = isDark ? 'Light Mode' : 'Dark Mode';
        }
    }

    if (themeToggle) themeToggle.addEventListener('click', toggle);
    if (themeToggleMobile) themeToggleMobile.addEventListener('click', toggle);

    // Listen for system theme changes
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
        if (!localStorage.getItem('doc2any_theme')) {
            if (e.matches) {
                document.documentElement.setAttribute('data-theme', 'dark');
            } else {
                document.documentElement.removeAttribute('data-theme');
            }
            updateThemeIcons(e.matches);
        }
    });
}

// ========================================
// Scroll Progress Bar
// ========================================

function initializeScrollProgress() {
    const progressBar = document.getElementById('scrollProgress');
    if (!progressBar) return;

    window.addEventListener('scroll', () => {
        const scrollTop = window.pageYOffset;
        const docHeight = document.documentElement.scrollHeight - window.innerHeight;
        const scrollPercent = docHeight > 0 ? (scrollTop / docHeight) * 100 : 0;
        progressBar.style.width = scrollPercent + '%';
    });
}

// ========================================
// Back to Top Button
// ========================================

function initializeBackToTop() {
    const btn = document.getElementById('backToTop');
    if (!btn) return;

    window.addEventListener('scroll', () => {
        if (window.pageYOffset > 400) {
            btn.classList.add('visible');
        } else {
            btn.classList.remove('visible');
        }
    });

    btn.addEventListener('click', () => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
}

// ========================================
// Typing Animation
// ========================================

function initializeTypingAnimation() {
    const el = document.getElementById('heroTitle');
    if (!el) return;

    const phrases = ['File Converter', 'PDF Converter', 'Image Converter', 'OCR Scanner', 'Doc2Any'];
    let phraseIndex = 0;
    let charIndex = 0;
    let isDeleting = false;

    // Add cursor
    const cursor = document.createElement('span');
    cursor.className = 'typing-cursor';
    el.appendChild(cursor);

    function type() {
        const current = phrases[phraseIndex];

        if (!isDeleting) {
            el.textContent = current.substring(0, charIndex + 1);
            el.appendChild(cursor);
            charIndex++;

            if (charIndex === current.length) {
                isDeleting = true;
                setTimeout(type, 2000);
                return;
            }
            setTimeout(type, 100);
        } else {
            el.textContent = current.substring(0, charIndex - 1);
            el.appendChild(cursor);
            charIndex--;

            if (charIndex === 0) {
                isDeleting = false;
                phraseIndex = (phraseIndex + 1) % phrases.length;
                setTimeout(type, 500);
                return;
            }
            setTimeout(type, 50);
        }
    }

    type();
}

// ========================================
// Mobile Menu
// ========================================

function initializeMobileMenu() {
    const menuBtn = document.getElementById('mobileMenuBtn');
    const menu = document.getElementById('mobileMenu');
    const closeBtn = document.getElementById('mobileMenuClose');

    if (!menuBtn || !menu) return;

    menuBtn.addEventListener('click', () => {
        menu.classList.add('active');
        document.body.style.overflow = 'hidden';
    });

    function close() {
        menu.classList.remove('active');
        document.body.style.overflow = '';
    }

    if (closeBtn) closeBtn.addEventListener('click', close);

    // Close on overlay click
    menu.addEventListener('click', (e) => {
        if (e.target === menu) close();
    });

    // Close on link click
    menu.querySelectorAll('.mobile-nav-link').forEach(link => {
        link.addEventListener('click', close);
    });

    // Close on ESC
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && menu.classList.contains('active')) close();
    });
}

// ========================================
// Ripple Effect on Buttons
// ========================================

function initializeRippleEffect() {
    const selectors = '.btn-convert, .btn-download, .quick-btn, .format-btn, .btn-ocr-convert, .btn-ocr-new, .contact-option';

    document.addEventListener('click', (e) => {
        const target = e.target.closest(selectors);
        if (!target) return;

        target.style.position = target.style.position || 'relative';
        target.style.overflow = 'hidden';

        const ripple = document.createElement('span');
        ripple.className = 'ripple';
        const rect = target.getBoundingClientRect();
        const size = Math.max(rect.width, rect.height);
        ripple.style.width = ripple.style.height = size + 'px';
        ripple.style.left = (e.clientX - rect.left - size / 2) + 'px';
        ripple.style.top = (e.clientY - rect.top - size / 2) + 'px';

        target.appendChild(ripple);
        ripple.addEventListener('animationend', () => ripple.remove());
    });
}
