const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const fs = require('fs').promises;
const path = require('path');
const multer = require('multer');
const PDFLib = require('pdf-lib');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3010;
const DATA_FILE = path.join(__dirname, 'data', 'projects.json');

function calculateCostBreakdown(subcontractorTotal, project, changeOrderData = null) {
    const feePercentage = (project && project.feePercentage) ? project.feePercentage : 10;
    const bondPercentage = (project && project.bondPercentage) ? project.bondPercentage : 1.5;
    const isOfcc = !!(project && project.isOfcc);
    const roundCurrency = (amount) => Math.round((amount || 0) * 100) / 100;
    const roundedSubcontractorTotal = roundCurrency(subcontractorTotal);

    let calculatedFeeAmount;
    let calculatedBondAmount;

    if (isOfcc) {
        const initialFeeAmount = roundCurrency(roundedSubcontractorTotal * (feePercentage / 100));
        const bondBase = roundCurrency(roundedSubcontractorTotal + initialFeeAmount);
        calculatedBondAmount = roundCurrency(bondBase * (bondPercentage / 100));
        calculatedFeeAmount = roundCurrency((roundedSubcontractorTotal + calculatedBondAmount) * (feePercentage / 100));
    } else {
        calculatedFeeAmount = roundCurrency(roundedSubcontractorTotal * (feePercentage / 100));
        const bondBase = roundCurrency(roundedSubcontractorTotal + calculatedFeeAmount);
        calculatedBondAmount = roundCurrency(bondBase * (bondPercentage / 100));
    }

    const hasManualFee = changeOrderData && typeof changeOrderData.manualFeeAmount === 'number' && !Number.isNaN(changeOrderData.manualFeeAmount);
    const hasManualBond = changeOrderData && typeof changeOrderData.manualBondAmount === 'number' && !Number.isNaN(changeOrderData.manualBondAmount);
    const feeAmount = hasManualFee ? roundCurrency(changeOrderData.manualFeeAmount) : calculatedFeeAmount;
    const bondAmount = hasManualBond ? roundCurrency(changeOrderData.manualBondAmount) : calculatedBondAmount;
    const totalCost = roundCurrency(roundedSubcontractorTotal + feeAmount + bondAmount);

    return {
        feeAmount,
        bondAmount,
        subtotal: roundedSubcontractorTotal,
        totalCost,
        calculatedFeeAmount,
        calculatedBondAmount,
        hasManualFee,
        hasManualBond
    };
}

// Middleware
app.use(cors());
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));

// Serve static files from public directory
app.use(express.static('public'));

// Function to sanitize file/folder names
function sanitizeFileName(name) {
    if (!name) return 'unknown';
    return name
        .replace(/[^a-z0-9\-_.]/gi, '_') // Replace invalid characters with underscore
        .replace(/_{2,}/g, '_') // Replace multiple underscores with single underscore
        .replace(/^_+|_+$/g, '') // Remove leading/trailing underscores
        .substring(0, 50); // Limit length
}

// Configure multer to store files temporarily, then move them after we have the form data
const upload = multer({
    dest: path.join(__dirname, 'data', 'temp'), // Temporary storage
    limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
    fileFilter: function (req, file, cb) {
        console.log('File filter - mimetype:', file.mimetype, 'originalname:', file.originalname);
    
        // Allow common file types including Excel
        const allowedTypes = /jpeg|jpg|png|gif|pdf|doc|docx|xls|xlsx|txt/;
        const extname = allowedTypes.test(path.extname(file.originalname).toLowerCase());
        const mimetype = allowedTypes.test(file.mimetype) || 
                        file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                        file.mimetype === 'application/vnd.ms-excel';
    
        if (mimetype && extname) {
            console.log('File type allowed');
            return cb(null, true);
        } else {
            console.log('File type rejected');
            cb(new Error('Only images, PDFs, Excel, and office documents are allowed!'));
        }
    }
});

// Helper function to populate PDF form fields
async function populateChangeOrderForm(form, linkedCosts, changeOrderData, project) {
    try {
        console.log('Populating PDF form fields...');
        console.log('Form has fields:', form.getFields().length);
        
        // Calculate totals with proper rounding (same as Excel)
        const subcontractorTotal = Math.round(linkedCosts.reduce((sum, cost) => sum + (cost.amount || 0), 0) * 100) / 100;
        const { feeAmount, bondAmount, subtotal, totalCost } = calculateCostBreakdown(subcontractorTotal, project, changeOrderData);
        
        console.log('PDF Calculated values:');
        console.log('- Subcontractor Total:', subcontractorTotal);
        console.log('- Fee Amount:', feeAmount);
        console.log('- Bond Amount:', bondAmount);
        console.log('- Total Cost:', totalCost);
        
        // Helper function to safely set text field
        const setTextField = (fieldName, value) => {
            try {
                const field = form.getTextField(fieldName);
                if (field && value !== null && value !== undefined) {
                    field.setText(value.toString());
                    console.log(`Set field '${fieldName}' to '${value}'`);
                }
            } catch (err) {
                console.log(`Field '${fieldName}' not found or not a text field`);
            }
        };
        
        // Map change order basic data to form fields
        setTextField('CO Number', changeOrderData.number || changeOrderData.itemNumber || '');
        setTextField('Description', changeOrderData.description || '');
        setTextField('Date', changeOrderData.dateAdded || '');
        setTextField('Status', changeOrderData.status || '');
        
        // Set calculated totals with proper formatting
        setTextField('Subcontractor Total', `$${subcontractorTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        setTextField('Fee', `$${feeAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        setTextField('Bond', `$${bondAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        setTextField('Subtotal', `$${subtotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        setTextField('Total Cost', `$${totalCost.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        
        // Populate individual subcontractor data
        console.log('Populating subcontractor data for', linkedCosts.length, 'costs');
        linkedCosts.forEach((cost, index) => {
            const subNum = index + 1;
            console.log(`Processing subcontractor ${subNum}:`, cost.subcontractor);
            
            // Subcontractor name
            setTextField(`Subcontractor ${subNum}`, cost.subcontractor || '');
            
            // Work performed
            setTextField(`Subcontractor ${subNum} work`, cost.workPerformed || '');
            
            // Amount (formatted as currency)
            const formattedAmount = cost.amount ? 
                `$${cost.amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : 
                '';
            setTextField(`Subcontractor ${subNum} $`, formattedAmount);
            
            // Additional fields
            setTextField(`Subcontractor ${subNum} CO`, cost.submittedCONumber || '');
            setTextField(`Subcontractor ${subNum} Description`, cost.description || '');
        });
        
        // List all available form fields for debugging
        console.log('Available form fields:');
        form.getFields().forEach(field => {
            console.log(`- ${field.getName()} (${field.constructor.name})`);
        });
        
    } catch (error) {
        console.error('Error populating PDF form:', error);
        throw error;
    }
}

// Helper function to populate Excel template with calculations
async function populateExcelTemplate(worksheet, linkedCosts, changeOrderData, project) {
    try {
        console.log('Starting template population...');
        console.log('Worksheet defined:', !!worksheet);
        console.log('Project data:', project ? 'exists' : 'missing');
        
        if (!worksheet) {
            throw new Error('Worksheet is undefined');
        }
        
        // Calculate totals with proper rounding
        const subcontractorTotal = Math.round(linkedCosts.reduce((sum, cost) => sum + (cost.amount || 0), 0) * 100) / 100;
        const { feeAmount, bondAmount, subtotal, totalCost } = calculateCostBreakdown(subcontractorTotal, project, changeOrderData);
        
        console.log('Calculated values:');
        console.log('- Subcontractor Total:', subcontractorTotal);
        console.log('- Fee Amount:', feeAmount);
        console.log('- Subtotal:', subtotal);
        console.log('- Bond Amount:', bondAmount);
        console.log('- Total Cost:', totalCost);
        
        // Function to find and replace tagged cells
        const findAndReplaceInWorksheet = (searchText, replaceText) => {
            console.log(`Looking for: ${searchText}, replacing with: ${replaceText}`);
            let found = false;
            
            if (typeof worksheet.eachRow !== 'function') {
                console.error('Worksheet does not have eachRow method');
                return;
            }
            
            try {
                worksheet.eachRow((row, rowNumber) => {
                    if (row && typeof row.eachCell === 'function') {
                        row.eachCell((cell, colNumber) => {
                            if (cell && cell.value && typeof cell.value === 'string' && cell.value.includes(searchText)) {
                                console.log(`Found ${searchText} at row ${rowNumber}, col ${colNumber}`);
                                cell.value = cell.value.replace(searchText, replaceText || '');
                                found = true;
                            }
                        });
                    }
                });
                if (!found) {
                    console.log(`Tag ${searchText} not found in worksheet`);
                }
            } catch (err) {
                console.error(`Error processing ${searchText}:`, err.message);
            }
        };
        
        // Populate basic change order info
        findAndReplaceInWorksheet('{CO Number}', changeOrderData.number || changeOrderData.itemNumber || '');
        findAndReplaceInWorksheet('{Description}', changeOrderData.description || '');
        findAndReplaceInWorksheet('{Date}', changeOrderData.dateAdded || '');
        findAndReplaceInWorksheet('{Status}', changeOrderData.status || '');
        
        // Populate calculated totals with proper formatting
        findAndReplaceInWorksheet('{Subcontractor Total}', `$${subcontractorTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        findAndReplaceInWorksheet('{Fee}', `$${feeAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        findAndReplaceInWorksheet('{Bond}', `$${bondAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        findAndReplaceInWorksheet('{Subtotal}', `$${subtotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        findAndReplaceInWorksheet('{Total Cost}', `$${totalCost.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        
        // Populate individual subcontractor data
        console.log('Populating subcontractor data for', linkedCosts.length, 'costs');
        linkedCosts.forEach((cost, index) => {
            const subNum = index + 1;
            console.log(`Processing subcontractor ${subNum}:`, cost.subcontractor);
            
            // Subcontractor name
            findAndReplaceInWorksheet(`{Subcontractor ${subNum}}`, cost.subcontractor || '');
            
            // Work performed
            findAndReplaceInWorksheet(`{Subcontractor ${subNum} work}`, cost.workPerformed || '');
            
            // Amount (formatted as currency)
            const formattedAmount = cost.amount ? 
                `$${cost.amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : 
                '';
            findAndReplaceInWorksheet(`{Subcontractor ${subNum} $}`, formattedAmount);
            
            // Additional fields
            findAndReplaceInWorksheet(`{Subcontractor ${subNum} CO}`, cost.submittedCONumber || '');
            findAndReplaceInWorksheet(`{Subcontractor ${subNum} Description}`, cost.description || '');
        });
        
        // Clear any remaining unused subcontractor placeholders
        const maxSubcontractorsToCheck = 20; // Adjust this number based on your template
        const usedSubcontractors = linkedCosts.length;
        
        console.log(`Clearing unused subcontractor tags from ${usedSubcontractors + 1} to ${maxSubcontractorsToCheck}`);
        
        for (let i = usedSubcontractors + 1; i <= maxSubcontractorsToCheck; i++) {
            findAndReplaceInWorksheet(`{Subcontractor ${i}}`, '');
            findAndReplaceInWorksheet(`{Subcontractor ${i} work}`, '');
            findAndReplaceInWorksheet(`{Subcontractor ${i} $}`, '');
            findAndReplaceInWorksheet(`{Subcontractor ${i} CO}`, '');
            findAndReplaceInWorksheet(`{Subcontractor ${i} Description}`, '');
        }
        
        // Clear any other common unused tags that might exist
        const commonUnusedTags = [
            '{Project Name}',
            '{Project Number}',
            '{Contractor}',
            '{Date Prepared}',
            '{Prepared By}',
        ];
        
        commonUnusedTags.forEach(tag => {
            findAndReplaceInWorksheet(tag, '');
        });
        
        console.log('Template population completed');
        
    } catch (error) {
        console.error('Error in populateExcelTemplate:', error);
        throw error;
    }
}

async function prewarmLibreOffice() {
    try {
        console.log('🔥 Pre-warming LibreOffice...');
        const { exec } = require('child_process');
        const util = require('util');
        const execPromise = util.promisify(exec);
        
        const env = {
            ...process.env,
            DISPLAY: ':99',
            HOME: process.env.HOME || '/tmp',
            SAL_DISABLE_OPENCL: '1',
            SAL_NO_OOSPLASH: '1',
        };
        
        // Just get the version - this loads LibreOffice into memory
        await execPromise('timeout 45 libreoffice --headless --invisible --version', { 
            env,
            timeout: 45000 
        });
        
        console.log('✅ LibreOffice pre-warmed successfully');
    } catch (error) {
        console.log('⚠️ LibreOffice pre-warm failed (this is not critical):', error.message);
        console.log('   PDF generation may be slower on first use');
    }
}

// Ensure data directory exists
async function ensureDataDirectory() {
    try {
        await fs.mkdir(path.join(__dirname, 'data'), { recursive: true });
        await fs.mkdir(path.join(__dirname, 'data', 'temp'), { recursive: true });
        
        // Check if projects file exists, if not create it with empty array
        try {
            await fs.access(DATA_FILE);
        } catch {
            await fs.writeFile(DATA_FILE, JSON.stringify({ projects: [] }, null, 2));
        }
    } catch (error) {
        console.error('Error creating data directory:', error);
    }
}

// Load projects from file
async function loadProjects() {
    try {
        const data = await fs.readFile(DATA_FILE, 'utf8');
        return JSON.parse(data);
    } catch (error) {
        console.error('Error loading projects:', error);
        return { projects: [] };
    }
}

// Save projects to file
async function saveProjects(data) {
    try {
        await fs.writeFile(DATA_FILE, JSON.stringify(data, null, 2));
        return true;
    } catch (error) {
        console.error('Error saving projects:', error);
        return false;
    }
}

// API Routes

// Template upload route - accepts both Excel and PDF files
app.post('/api/upload-template', upload.single('template'), async (req, res) => {
    try {
        const { projectId } = req.body;
        
        if (!req.file) {
            return res.status(400).json({ error: 'No template file uploaded' });
        }
        
        // Check if it's an Excel file or PDF
        const allowedExcelTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel' // .xls
        ];
        
        const allowedPdfTypes = [
            'application/pdf'
        ];
        
        const isExcel = allowedExcelTypes.includes(req.file.mimetype);
        const isPdf = allowedPdfTypes.includes(req.file.mimetype);
        
        if (!isExcel && !isPdf) {
            return res.status(400).json({ error: 'Please upload an Excel file (.xlsx or .xls) or PDF file (.pdf)' });
        }
        
        // Move template to project-specific location
        const templatePath = path.join(__dirname, 'data', 'templates');
        await fs.mkdir(templatePath, { recursive: true });
        
        const fileExtension = path.extname(req.file.originalname);
        const finalPath = path.join(templatePath, `project_${projectId}_template${fileExtension}`);
        await fs.rename(req.file.path, finalPath);
        
        res.json({ 
            success: true, 
            templatePath: path.relative(path.join(__dirname, 'data'), finalPath).replace(/\\/g, '/'),
            message: `${isExcel ? 'Excel' : 'PDF'} template uploaded successfully` 
        });
    } catch (error) {
        console.error('Template upload error:', error);
        res.status(500).json({ error: 'Failed to upload template' });
    }
});

// Excel-based generation route
app.post('/api/generate-change-order-excel', async (req, res) => {
    try {
        const { projectId, changeOrderId, linkedCosts, changeOrderData, project } = req.body;
        
        console.log('=== EXCEL GENERATION ===');
        console.log('Project ID:', projectId);
        console.log('Linked costs count:', linkedCosts ? linkedCosts.length : 0);
        
        // Find the template file (could be .xlsx or .xls)
        const templateDir = path.join(__dirname, 'data', 'templates');
        const possibleExtensions = ['.xlsx', '.xls'];
        let templatePath = null;
        
        for (const ext of possibleExtensions) {
            const testPath = path.join(templateDir, `project_${projectId}_template${ext}`);
            try {
                await fs.access(testPath);
                templatePath = testPath;
                console.log('Found template:', templatePath);
                break;
            } catch (err) {
                console.log('Template not found at:', testPath);
            }
        }
        
        if (!templatePath) {
            return res.status(404).json({ error: 'Excel template not found. Please upload an Excel template first.' });
        }
        
        // Load the Excel workbook
        const workbook = new ExcelJS.Workbook();
        console.log('Loading Excel file...');
        await workbook.xlsx.readFile(templatePath);
        
        // Get the first worksheet
        console.log('Worksheet count:', workbook.worksheets.length);
        const worksheet = workbook.worksheets[0];
        
        if (!worksheet) {
            return res.status(500).json({ error: 'No worksheets found in the Excel template' });
        }
        
        console.log('Worksheet name:', worksheet.name);
        console.log('Worksheet has rows:', worksheet.rowCount);
        
        // Populate the Excel template
        console.log('Populating template...');
        await populateExcelTemplate(worksheet, linkedCosts, changeOrderData, project);
        
        // Save the filled Excel file temporarily
        const tempExcelPath = path.join(__dirname, 'data', 'temp', `CO_${changeOrderData.number || changeOrderData.itemNumber}_${Date.now()}.xlsx`);
        console.log('Saving to temp path:', tempExcelPath);
        await workbook.xlsx.writeFile(tempExcelPath);
        
        // Read the file back
        const excelBuffer = await fs.readFile(tempExcelPath);
        
        // Clean up temp file
        await fs.unlink(tempExcelPath);
        
        // Send as downloadable Excel file
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="CO_${changeOrderData.number || changeOrderData.itemNumber}_${Date.now()}.xlsx"`);
        res.send(excelBuffer);
        
    } catch (error) {
        console.error('Excel generation error:', error);
        res.status(500).json({ error: 'Failed to generate Excel file: ' + error.message });
    }
});

// Enhanced PDF generation route with better error handling and LibreOffice optimization
app.post('/api/generate-change-order-pdf', async (req, res) => {
    try {
        const { projectId, changeOrderId, linkedCosts, changeOrderData, project } = req.body;
        
        console.log('=== PDF GENERATION (Excel->PDF + Merge) ===');
        console.log('Project ID:', projectId);
        console.log('Linked costs count:', linkedCosts ? linkedCosts.length : 0);
        
        // Find the Excel template file (same as Excel generation)
        const templateDir = path.join(__dirname, 'data', 'templates');
        const possibleExtensions = ['.xlsx', '.xls'];
        let templatePath = null;
        
        for (const ext of possibleExtensions) {
            const testPath = path.join(templateDir, `project_${projectId}_template${ext}`);
            try {
                await fs.access(testPath);
                templatePath = testPath;
                console.log('Found Excel template:', templatePath);
                break;
            } catch (err) {
                console.log('Template not found at:', testPath);
            }
        }
        
        if (!templatePath) {
            return res.status(404).json({ error: 'Excel template not found. Please upload an Excel template first.' });
        }
        
        // Step 1: Load and populate the Excel workbook
        const workbook = new ExcelJS.Workbook();
        console.log('Loading Excel file...');
        await workbook.xlsx.readFile(templatePath);
        
        const worksheet = workbook.worksheets[0];
        if (!worksheet) {
            return res.status(500).json({ error: 'No worksheets found in the Excel template' });
        }
        
        console.log('Populating Excel template for PDF conversion...');
        await populateExcelTemplate(worksheet, linkedCosts, changeOrderData, project);
        
        // Step 2: Save the populated Excel file temporarily
        const timestamp = Date.now();
        const tempExcelPath = path.join(__dirname, 'data', 'temp', `CO_${changeOrderData.number || changeOrderData.itemNumber}_${timestamp}.xlsx`);
        console.log('Saving populated Excel to:', tempExcelPath);
        await workbook.xlsx.writeFile(tempExcelPath);
        
        // Step 3: Convert Excel to PDF with enhanced LibreOffice handling
        const { exec } = require('child_process');
        const util = require('util');
        const execPromise = util.promisify(exec);
        
        const tempPdfPath = tempExcelPath.replace('.xlsx', '.pdf').replace('.xls', '.pdf');
        
        try {
            console.log('Converting Excel to PDF using LibreOffice...');
            
            // Enhanced LibreOffice command for Raspberry Pi
            const tempDir = path.dirname(tempPdfPath);
            const command = `timeout 120 libreoffice --headless --invisible --nodefault --nolockcheck --nologo --norestore --convert-to pdf --outdir "${tempDir}" "${tempExcelPath}"`;
            console.log('LibreOffice command:', command);
            
            // Set environment variables for better headless operation
            const env = {
                ...process.env,
                DISPLAY: ':99',
                HOME: process.env.HOME || '/tmp',
                TMPDIR: tempDir,
                SAL_DISABLE_OPENCL: '1',
                SAL_NO_OOSPLASH: '1',
            };
            
            const { stdout, stderr } = await execPromise(command, { 
                timeout: 120000, // 2 minutes timeout
                env: env,
                maxBuffer: 1024 * 1024 * 10
            });
            
            console.log('LibreOffice stdout:', stdout);
            if (stderr) console.log('LibreOffice stderr:', stderr);
            
            // Wait for file system sync on Raspberry Pi
            await new Promise(resolve => setTimeout(resolve, 2000));
            
            // Check if PDF was created with retries
            let pdfExists = false;
            for (let i = 0; i < 5; i++) {
                try {
                    await fs.access(tempPdfPath);
                    pdfExists = true;
                    console.log('PDF conversion successful:', tempPdfPath);
                    break;
                } catch (accessError) {
                    console.log(`PDF check attempt ${i + 1}/5 failed, waiting...`);
                    await new Promise(resolve => setTimeout(resolve, 1000));
                }
            }
            
            if (!pdfExists) {
                throw new Error(`PDF file not created at ${tempPdfPath}. LibreOffice conversion may have failed.`);
            }
            
            // Load the converted PDF
            const convertedPdfBytes = await fs.readFile(tempPdfPath);
            const pdfDoc = await PDFLib.PDFDocument.load(convertedPdfBytes);
            console.log('Successfully loaded converted PDF with', pdfDoc.getPageCount(), 'pages');
            
            // Step 4: Merge attachment files (your existing code continues here...)
            console.log('Merging attachment files...');
            for (const cost of linkedCosts) {
                if (cost.selectedFile) {
                    const filePath = path.join(__dirname, 'data', cost.selectedFile);
                    try {
                        const fileBuffer = await fs.readFile(filePath);
                        const fileExt = path.extname(cost.selectedFile).toLowerCase();
                        
                        if (fileExt === '.pdf') {
                            console.log('Merging PDF attachment:', cost.selectedFile);
                            const attachmentPdf = await PDFLib.PDFDocument.load(fileBuffer);
                            const copiedPages = await pdfDoc.copyPages(attachmentPdf, attachmentPdf.getPageIndices());
                            copiedPages.forEach((page) => pdfDoc.addPage(page));
                        } else {
                            console.log('Adding non-PDF attachment info:', cost.selectedFile);
                            const attachmentPage = pdfDoc.addPage();
                            const { height: attachmentHeight } = attachmentPage.getSize();
                            
                            attachmentPage.drawText(`Attachment: ${cost.originalName || cost.selectedFile}`, {
                                x: 50, y: attachmentHeight - 50, size: 14,
                            });
                            attachmentPage.drawText(`From: ${cost.subcontractor}`, {
                                x: 50, y: attachmentHeight - 80, size: 12,
                            });
                            attachmentPage.drawText(`Description: ${cost.description}`, {
                                x: 50, y: attachmentHeight - 110, size: 12,
                            });
                            attachmentPage.drawText(`Amount: $${(cost.amount || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`, {
                                x: 50, y: attachmentHeight - 140, size: 12,
                            });
                            attachmentPage.drawText('Note: Original file was attached to this change order', {
                                x: 50, y: attachmentHeight - 200, size: 10,
                            });
                            attachmentPage.drawText('but cannot be embedded in PDF format.', {
                                x: 50, y: attachmentHeight - 220, size: 10,
                            });
                        }
                    } catch (fileError) {
                        console.error('Error processing attachment file:', fileError);
                    }
                }
            }
            
            // Clean up temp files
            try {
                await fs.unlink(tempExcelPath);
                await fs.unlink(tempPdfPath);
            } catch (cleanupError) {
                console.error('Error cleaning up temp files:', cleanupError);
            }
            
            // Generate final PDF
            const pdfBytes = await pdfDoc.save();
            
            console.log('PDF generation completed successfully');
            
            // Send as downloadable PDF
            res.setHeader('Content-Type', 'application/pdf');
            res.setHeader('Content-Disposition', `attachment; filename="CO_${changeOrderData.number || changeOrderData.itemNumber}_${timestamp}.pdf"`);
            res.send(Buffer.from(pdfBytes));
            
        } catch (conversionError) {
            console.error('LibreOffice conversion failed:', conversionError);
            
            // Clean up temp Excel file
            try {
                await fs.unlink(tempExcelPath);
            } catch (cleanupError) {
                console.error('Error cleaning up temp Excel file:', cleanupError);
            }
            
            // Provide helpful error messages
            if (conversionError.signal === 'SIGTERM' || conversionError.code === null) {
                return res.status(500).json({ 
                    error: 'LibreOffice conversion timed out. This is common on Raspberry Pi. Please try the Excel option instead, or run the setup script to configure LibreOffice properly.',
                    details: 'LibreOffice timeout - try Excel generation or setup LibreOffice'
                });
            } else {
                return res.status(500).json({ 
                    error: 'Failed to convert Excel to PDF. Please run the LibreOffice setup script or use the Excel option instead.',
                    details: conversionError.message
                });
            }
        }
        
    } catch (error) {
        console.error('PDF generation error:', error);
        res.status(500).json({ error: 'Failed to generate PDF file: ' + error.message });
    }
});
// Get all projects
app.get('/api/projects', async (req, res) => {
    try {
        const data = await loadProjects();
        res.json(data.projects);
    } catch (error) {
        res.status(500).json({ error: 'Failed to load projects' });
    }
});

// Save all projects
app.post('/api/projects', async (req, res) => {
    try {
        const { projects } = req.body;
        const success = await saveProjects({ 
            projects,
            lastUpdated: new Date().toISOString()
        });
        
        if (success) {
            res.json({ success: true, message: 'Projects saved successfully' });
        } else {
            res.status(500).json({ error: 'Failed to save projects' });
        }
    } catch (error) {
        res.status(500).json({ error: 'Failed to save projects' });
    }
});

// Import projects
app.post('/api/import', async (req, res) => {
    try {
        const { projects, mode } = req.body; // mode: 'merge' or 'replace'
        
        let currentData = { projects: [] };
        if (mode === 'merge') {
            currentData = await loadProjects();
        }
        
        // For merge mode, assign new IDs to avoid conflicts
        if (mode === 'merge') {
            const maxExistingId = currentData.projects.length > 0 ? 
                Math.max(...currentData.projects.map(p => p.id)) : 0;
            
            const importedProjects = projects.map((project, index) => ({
                ...project,
                id: maxExistingId + index + 1
            }));
            
            currentData.projects = [...currentData.projects, ...importedProjects];
        } else {
            currentData.projects = projects;
        }
        
        const success = await saveProjects({
            ...currentData,
            lastUpdated: new Date().toISOString()
        });
        
        if (success) {
            res.json({ 
                success: true, 
                message: `Successfully imported ${projects.length} projects`,
                count: currentData.projects.length
            });
        } else {
            res.status(500).json({ error: 'Failed to import projects' });
        }
    } catch (error) {
        res.status(500).json({ error: 'Failed to import projects' });
    }
});

// Export projects
app.get('/api/export', async (req, res) => {
    try {
        const data = await loadProjects();
        const exportData = {
            version: "1.0",
            exportDate: new Date().toISOString(),
            projects: data.projects
        };
        
        res.setHeader('Content-Type', 'application/json');
        res.setHeader('Content-Disposition', 'attachment; filename="construction_projects_export.json"');
        res.json(exportData);
    } catch (error) {
        res.status(500).json({ error: 'Failed to export projects' });
    }
});

// Enhanced file upload endpoint with organized storage
app.post('/api/upload-cost-files', upload.array('files'), async (req, res) => {
    try {
        console.log('=== UPLOAD ENDPOINT CALLED ===');
        console.log('req.body:', req.body);
        console.log('req.files length:', req.files ? req.files.length : 0);
        
        const { projectId, projectName, subcontractor, costId, description } = req.body;
        
        console.log('Form data received:');
        console.log('- projectId:', projectId);
        console.log('- projectName:', projectName);
        console.log('- subcontractor:', subcontractor);
        console.log('- costId:', costId);
        console.log('- description:', description);
        
        if (!req.files || req.files.length === 0) {
            console.log('ERROR: No files uploaded');
            return res.status(400).json({ error: 'No files uploaded' });
        }
        
        if (!projectName || !subcontractor || !costId) {
            console.log('ERROR: Missing required fields');
            return res.status(400).json({ 
                error: 'Missing required fields: projectName, subcontractor, or costId',
                received: { projectName, subcontractor, costId }
            });
        }
        
        // Create sanitized folder names
        const sanitizedProjectName = sanitizeFileName(projectName);
        const sanitizedSubcontractor = sanitizeFileName(subcontractor);
        const sanitizedCostId = sanitizeFileName(costId.toString());
        
        console.log('Sanitized values:');
        console.log('- sanitizedProjectName:', sanitizedProjectName);
        console.log('- sanitizedSubcontractor:', sanitizedSubcontractor);
        console.log('- sanitizedCostId:', sanitizedCostId);
        
        // Create the organized folder path: data/projectName/subcontractor/costId
        const finalPath = path.join(__dirname, 'data', sanitizedProjectName, sanitizedSubcontractor, sanitizedCostId);
        console.log('Final path:', finalPath);
        
        // Ensure the directory exists
        await fs.mkdir(finalPath, { recursive: true });
        console.log('Directory created successfully');
        
        // Move files from temp to final location and process them
        const uploadedFiles = [];
        
        for (const file of req.files) {
            const timestamp = Date.now();
            const sanitizedOriginalName = sanitizeFileName(file.originalname);
            const finalFilename = `${timestamp}-${sanitizedOriginalName}`;
            const finalFilePath = path.join(finalPath, finalFilename);
            
            // Move file from temp to final location
            await fs.rename(file.path, finalFilePath);
            
            const fileInfo = {
                filename: finalFilename,
                originalName: file.originalname,
                size: file.size,
                mimetype: file.mimetype,
                uploadDate: new Date().toISOString(),
                relativePath: path.relative(path.join(__dirname, 'data'), finalFilePath).replace(/\\/g, '/'),
                fullPath: finalFilePath
            };
            
            uploadedFiles.push(fileInfo);
            console.log('Processed file:', finalFilename);
        }
        
        console.log('Successfully processed files:', uploadedFiles.map(f => f.filename));
        console.log(`Uploaded ${uploadedFiles.length} files for project: ${projectName}, subcontractor: ${subcontractor}, cost: ${costId}`);
        
        res.json({ 
            success: true, 
            files: uploadedFiles,
            message: `Successfully uploaded ${uploadedFiles.length} file(s)`,
            debug: {
                projectName,
                subcontractor,
                costId,
                finalPath
            }
        });
    } catch (error) {
        console.error('File upload error:', error);
        res.status(500).json({ error: 'Failed to upload files: ' + error.message });
    }
});

// Get files for a specific cost
app.get('/api/cost-files/:projectName/:subcontractor/:costId', async (req, res) => {
    try {
        const { projectName, subcontractor, costId } = req.params;
        
        console.log('=== GET FILES ENDPOINT ===');
        console.log('Params received:');
        console.log('- projectName:', projectName);
        console.log('- subcontractor:', subcontractor);
        console.log('- costId:', costId);
        
        const sanitizedProjectName = sanitizeFileName(projectName);
        const sanitizedSubcontractor = sanitizeFileName(subcontractor);
        const sanitizedCostId = sanitizeFileName(costId);
        
        console.log('Sanitized params:');
        console.log('- sanitizedProjectName:', sanitizedProjectName);
        console.log('- sanitizedSubcontractor:', sanitizedSubcontractor);
        console.log('- sanitizedCostId:', sanitizedCostId);
        
        const filesPath = path.join(__dirname, 'data', sanitizedProjectName, sanitizedSubcontractor, sanitizedCostId);
        console.log('Looking for files in:', filesPath);
        
        try {
            const files = await fs.readdir(filesPath);
            console.log('Found files:', files);
            
            const fileDetails = await Promise.all(
                files.map(async (filename) => {
                    const filePath = path.join(filesPath, filename);
                    const stats = await fs.stat(filePath);
                    return {
                        filename,
                        originalName: filename.substring(filename.indexOf('-') + 1), // Remove timestamp prefix
                        size: stats.size,
                        uploadDate: stats.mtime.toISOString(),
                        relativePath: path.relative(path.join(__dirname, 'data'), filePath).replace(/\\/g, '/')
                    };
                })
            );
            
            console.log('Returning file details:', fileDetails);
            res.json({ files: fileDetails });
        } catch (error) {
            console.log('Directory does not exist or is empty:', error.message);
            // Directory doesn't exist, return empty array
            res.json({ files: [] });
        }
    } catch (error) {
        console.error('Error getting files:', error);
        res.status(500).json({ error: 'Failed to get files' });
    }
});

// Serve uploaded files
app.use('/data', express.static(path.join(__dirname, 'data')));

// Delete a specific file
app.delete('/api/delete-file', async (req, res) => {
    try {
        const { relativePath } = req.body;
        console.log('Deleting file:', relativePath);
        
        const fullPath = path.join(__dirname, 'data', relativePath);
        
        // Security check: ensure the path is within the data directory
        const resolvedPath = path.resolve(fullPath);
        const dataPath = path.resolve(path.join(__dirname, 'data'));
        
        if (!resolvedPath.startsWith(dataPath)) {
            return res.status(400).json({ error: 'Invalid file path' });
        }
        
        await fs.unlink(resolvedPath);
        console.log('File deleted successfully:', relativePath);
        res.json({ success: true, message: 'File deleted successfully' });
    } catch (error) {
        console.error('Error deleting file:', error);
        res.status(500).json({ error: 'Failed to delete file' });
    }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'ok', 
        timestamp: new Date().toISOString(),
        server: 'Cost Tracker Server'
    });
});

// Serve the main application
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Start server
async function startServer() {
    await ensureDataDirectory();
    
    // Add this line to prewarm LibreOffice
    await prewarmLibreOffice();
    
    app.listen(PORT, '0.0.0.0', () => {
        console.log(`Cost Tracker Server running on http://0.0.0.0:3010`);
        console.log(`Access from network at http://10.0.10.180:3010`);
        console.log(`Data stored in: ${DATA_FILE}`);
    });
}

startServer().catch(console.error);
